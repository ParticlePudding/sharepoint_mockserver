"""SharePoint REST API の簡易モックサーバー（座席表・会議室予約の開発用）

1つのサーバーで、apps.json に書いた複数の Web アプリを、本物の SharePoint と同じ形の URL で配信する。
  http://localhost:8000/                                   … アプリの一覧
  http://localhost:8000/SiteAssets/<name>/index.html        … 各アプリ（このサーバーのデータを使う）
  http://localhost:8000/_api/...                           … SharePoint REST API のまね（全アプリ共通）

使い方:
    python mockserver.py                 # http://localhost:8000/ で起動（この PC からだけ開ける）
    python mockserver.py --reset         # DB を消してサンプルデータを入れ直す
    python mockserver.py --port 8080     # ポートを変える
    python mockserver.py --config 別の.json
    python mockserver.py --host 0.0.0.0  # 同じネットワークのスマホなどからも開けるようにする（ポートの許可が必要）

※ 各アプリは URL に ?mock=1 がある時だけ「ブラウザ内のテストデータ」で動く（保存されない）。
  ?mock=1 を付けずに開けば、このサーバーのデータを使う。

対応している SharePoint REST の機能:
  - GET  /_api/web/lists/getbytitle('リスト名')/items[(ID)]
         $filter（eq/ne/gt/ge/lt/le、and/or、括弧） / $orderby / $top / $select
         Accept の odata=nometadata / minimalmetadata / verbose
         minimalmetadata 以上では各行に "odata.etag" を付与。単一アイテムは ETag ヘッダーも返す
  - POST /_api/contextinfo  （FormDigest の発行）
  - POST .../items                         新規追加（201）
  - POST .../items(ID) + X-HTTP-Method: MERGE / DELETE
         IF-Match が "*" 以外で現在の ETag と違えば 412（楽観ロック）
  - apps.json の uniqueTitleLists に書いたリストは、Title の重複を 409 で拒否（「一意の値を適用」のまね）
  - リストは最初に書き込んだ時点で自動的にできる（作る手順は不要）
"""

import argparse
import functools
import html
import http.server
import json
import re
import socket
import sqlite3
import threading
import urllib.parse
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_PORT = 8000
DEFAULT_CONFIG = SCRIPT_DIR / "apps.json"
DEFAULT_DB = SCRIPT_DIR / "mock_sharepoint.db"
DIGEST_VALUE = "MOCK_DIGEST_TOKEN_12345"
DIGEST_TIMEOUT_SECONDS = 1800
SP_DEFAULT_TOP = 100  # SharePoint は $top 未指定時に100件までしか返さない
JST = timezone(timedelta(hours=9))

DB_LOCK = threading.Lock()  # 読み取り→ETag確認→書き込み をまとめて排他する


# ===================== 設定（apps.json） =====================

def load_config(path):
    cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    base = Path(path).resolve().parent
    apps = []
    for a in cfg.get("apps", []):
        name = str(a["name"])
        if not re.fullmatch(r"[A-Za-z0-9_-]+", name):
            raise SystemExit(f"apps.json: name は英数字・-・_ だけにしてください: {name}")
        root = (base / a["root"]).resolve()
        apps.append({"name": name, "title": a.get("title", name), "root": root, "seed": a.get("seed")})
    return {"apps": apps, "unique": set(cfg.get("uniqueTitleLists", []))}


# ===================== sample-data.js の読み込み（座席表） =====================

def js_literal_to_json(src):
    """JSのオブジェクト/配列リテラル（キーにクォートなし・'文字列'・コメント・末尾カンマ）をJSONに変換する"""
    out = []
    i, n = 0, len(src)
    while i < n:
        c = src[i]
        if c in "\"'":
            quote, j, buf = c, i + 1, []
            while j < n and src[j] != quote:
                if src[j] == "\\" and j + 1 < n:
                    nxt = src[j + 1]
                    buf.append(nxt if nxt in "'\"" else "\\" + nxt)
                    j += 2
                    continue
                buf.append(src[j])
                j += 1
            raw = "".join(buf)
            out.append(json.dumps(json.loads('"' + raw.replace('"', '\\"') + '"'), ensure_ascii=False))
            i = j + 1
        elif src.startswith("//", i):
            while i < n and src[i] != "\n":
                i += 1
        elif src.startswith("/*", i):
            end = src.find("*/", i + 2)
            i = n if end < 0 else end + 2
        elif c.isalpha() or c in "_$":
            j = i
            while j < n and (src[j].isalnum() or src[j] in "_$"):
                j += 1
            word = src[i:j]
            k = j
            while k < n and src[k] in " \t\r\n":
                k += 1
            out.append(json.dumps(word) if k < n and src[k] == ":" else word)
            i = j
        else:
            out.append(c)
            i += 1
    text = re.sub(r",(\s*[\]}])", r"\1", "".join(out))  # 末尾カンマ
    return json.loads(text)


def extract_js_const(src, name):
    m = re.search(r"\b(?:const|let|var)\s+" + re.escape(name) + r"\s*=\s*", src)
    if not m:
        return None
    start = m.end()
    if start >= len(src) or src[start] not in "[{":
        return None
    depth, i, quote = 0, start, None
    while i < len(src):
        c = src[i]
        if quote:
            if c == "\\":
                i += 2
                continue
            if c == quote:
                quote = None
        elif c in "\"'":
            quote = c
        elif src.startswith("//", i):
            i = src.find("\n", i)
            if i < 0:
                break
            continue
        elif c in "[{":
            depth += 1
        elif c in "]}":
            depth -= 1
            if depth == 0:
                return js_literal_to_json(src[start:i + 1])
        i += 1
    return None


# ===================== サンプルデータ =====================

def seed_seatmap(root):
    """座席表: js/sample-data.js の SAMPLE_ITEMS / SAMPLE_STATS_GROUPS を読み込む"""
    sample_js = root / "js" / "sample-data.js"
    config_js = root / "js" / "config.js"
    if not sample_js.exists():
        print(f"[seed] {sample_js} がないため、座席表のサンプルデータは入れません")
        return {}
    src = sample_js.read_text(encoding="utf-8")
    items = extract_js_const(src, "SAMPLE_ITEMS") or []
    groups = extract_js_const(src, "SAMPLE_STATS_GROUPS")
    quick_notes = []
    if config_js.exists():
        m = re.search(r"DEFAULT_QUICK_NOTES\s*:\s*(\[[^\]]*\])", config_js.read_text(encoding="utf-8"))
        if m:
            quick_notes = js_literal_to_json(m.group(1))
    for item in items:
        if not item.get("QuickNotes"):
            item["QuickNotes"] = json.dumps(quick_notes, ensure_ascii=False)
    lists = {"社員名簿": items}
    if groups is not None:
        lists["AppSettings"] = [{"Id": 1, "Title": "StatsGroups", "Value": json.dumps(groups, ensure_ascii=False)}]
    return lists


def seed_roombook(_root):
    """会議室予約: 今日を基準に、架空の予約を作る（js/sample-data.js と同じ内容の簡易版）"""
    today = datetime.now(JST).date()
    rows = {}
    seq = [0]
    now_iso = datetime.now(timezone.utc).isoformat()

    def add(room, d, start, end, title, by, series=None):
        key = f"{room}_{d.isoformat()}"
        row = rows.setdefault(key, {"Title": key, "RoomId": room, "BookingDate": d.isoformat(), "list": []})
        seq[0] += 1
        b = {"id": f"seed{seq[0]}", "start": start, "end": end, "title": title, "by": by, "note": "",
             "createdAt": now_iso, "updatedAt": now_iso}
        if series:
            b["series"] = series
        row["list"].append(b)

    def series_dates(first, count, match):
        out, d = [], first
        while len(out) < count:
            if match(d):
                out.append(d)
            d += timedelta(days=1)
        return out

    first = today - timedelta(days=7)
    # JS の曜日は 0=日〜6=土、Python の weekday() は 0=月〜6=日
    js_wd = lambda d: (d.weekday() + 1) % 7
    defs = [
        ("R101", "09:00", "09:15", "営業部 朝会", "中野 玲奈", {"freq": "weekdays", "endType": "count", "count": 40},
         "平日（月〜金）（40回）", lambda d: d.weekday() < 5, first),
        ("R301", "10:00", "11:30", "部長会", "坂本 颯太",
         {"freq": "weekly", "interval": 1, "weekdays": [2], "endType": "count", "count": 12},
         "毎週 火（12回）", lambda d: js_wd(d) == 2, first),
    ]
    first_wed = first
    while not (js_wd(first_wed) == 3 and (first_wed.day + 6) // 7 == 1):
        first_wed += timedelta(days=1)
    defs.append(("R201", "15:00", "16:00", "月次報告会", "斎藤 千夏",
                 {"freq": "monthlyNth", "nth": 1, "endType": "count", "count": 6},
                 "毎月 第1水曜日（6回）", lambda d: js_wd(d) == 3 and (d.day + 6) // 7 == 1, first_wed))
    for i, (room, s, e, title, by, rule, text, match, start) in enumerate(defs, 1):
        dates = series_dates(start, rule["count"], match)
        info = {"id": f"seedseries{i}", "rule": rule, "start": dates[0].isoformat(), "text": text}
        for d in dates:
            add(room, d, s, e, title, by, info)

    singles = [
        (0, "R101", "10:00", "11:00", "A社 打合せ", "遠藤 修平"),
        (0, "R101", "13:30", "14:45", "採用面接", "後藤 若菜"),
        (0, "R102", "09:30", "10:20", "設計レビュー", "石川 沙織"),
        (0, "R102", "16:00", "17:30", "新製品 企画会議", "安藤 芽衣"),
        (0, "R103", "11:00", "11:45", "B社 来訪", "西村 亮介"),
        (0, "R202", "13:00", "14:00", "Web会議（大阪支社）", "岡本 弘美"),
        (0, "R203", "14:10", "14:40", "1on1", "中島 紗希"),
        (0, "R301", "13:00", "17:00", "新人研修", "後藤 若菜"),
        (0, "R302", "08:30", "09:00", "障害振り返り", "山崎 麻衣"),
        (0, "R303", "10:00", "12:00", "検証作業", "小島 涼"),
        (1, "R101", "15:00", "16:00", "C社 定例", "遠藤 修平"),
        (1, "R201", "10:00", "12:00", "予算会議", "坂本 颯太"),
        (1, "R104", "14:00", "15:00", "D社 来訪", "前田 綾香"),
        (2, "R102", "13:00", "15:00", "仕様打合せ", "石川 沙織"),
        (2, "R302", "09:00", "10:00", "品質会議", "高田 彩"),
        (3, "R202", "14:00", "17:00", "社内勉強会", "小島 涼"),
        (4, "R101", "11:00", "12:00", "E社 打合せ", "西村 亮介"),
        (6, "R301", "13:00", "18:00", "全社説明会", "斎藤 千夏"),
    ]
    for offset, room, s, e, title, by in singles:
        add(room, today + timedelta(days=offset), s, e, title, by)

    items = []
    for i, row in enumerate(sorted(rows.values(), key=lambda r: r["Title"]), 1):
        row["list"].sort(key=lambda b: b["start"])
        items.append({"Id": i, "Title": row["Title"], "RoomId": row["RoomId"], "BookingDate": row["BookingDate"],
                      "Bookings": json.dumps(row["list"], ensure_ascii=False)})
    return {"会議室予約": items}


SEEDERS = {"seatmap": seed_seatmap, "roombook": seed_roombook}


# ===================== DB =====================

def connect(db_path):
    conn = sqlite3.connect(db_path, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db(db_path, apps, reset=False):
    with connect(db_path) as conn:
        if reset:
            conn.execute("DROP TABLE IF EXISTS sp_items")
            conn.execute("DROP TABLE IF EXISTS seeded")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sp_items (
                list_title TEXT NOT NULL,
                id         INTEGER NOT NULL,
                version    INTEGER NOT NULL DEFAULT 1,
                data       TEXT NOT NULL,
                PRIMARY KEY (list_title, id)
            )
        """)
        # サンプルを入れたアプリを記録（予約を全部消しても、起動し直すたびに入れ直さないように）
        conn.execute("CREATE TABLE IF NOT EXISTS seeded (app TEXT PRIMARY KEY)")
        for app in apps:
            seeder = SEEDERS.get(app["seed"] or "")
            if not seeder or conn.execute("SELECT 1 FROM seeded WHERE app=?", (app["name"],)).fetchone():
                continue
            lists = seeder(app["root"])
            for list_title, items in lists.items():
                if conn.execute("SELECT 1 FROM sp_items WHERE list_title=? LIMIT 1", (list_title,)).fetchone():
                    print(f"[seed] {list_title} にはすでにデータがあるため、入れません")
                    continue
                for item in items:
                    item = dict(item)
                    item_id = int(item.pop("Id"))
                    item.pop("ID", None)
                    conn.execute("INSERT INTO sp_items (list_title, id, version, data) VALUES (?, ?, 1, ?)",
                                 (list_title, item_id, json.dumps(item, ensure_ascii=False)))
                print(f"[seed] {app['title']}: {list_title} に {len(items)}件 を入れました")
            conn.execute("INSERT INTO seeded (app) VALUES (?)", (app["name"],))


def make_etag(version):
    return f'"{version}"'


def row_to_item(row):
    item_id, version, data = row
    return {"Id": item_id, "ID": item_id, **json.loads(data)}, version


# ===================== OData クエリ =====================

class ODataError(Exception):
    pass


TOKEN = re.compile(r"\s*(?:(\()|(\))|('(?:[^']|'')*')|(-?\d+(?:\.\d+)?)|([A-Za-z_][A-Za-z0-9_/]*))", re.UNICODE)


def tokenize(expr):
    tokens, pos = [], 0
    expr = expr.strip()
    while pos < len(expr):
        m = TOKEN.match(expr, pos)
        if not m or m.end() == pos:
            raise ODataError(f"未対応の $filter 式です: {expr[pos:]}")
        lp, rp, s, num, word = m.groups()
        if lp: tokens.append(("(", None))
        elif rp: tokens.append((")", None))
        elif s is not None: tokens.append(("lit", s[1:-1].replace("''", "'")))
        elif num is not None: tokens.append(("lit", float(num) if "." in num else int(num)))
        else:
            low = word.lower()
            if low in ("and", "or", "eq", "ne", "gt", "ge", "lt", "le"):
                tokens.append((low, None))
            elif low in ("true", "false", "null"):
                tokens.append(("lit", {"true": True, "false": False, "null": None}[low]))
            else:
                tokens.append(("field", word))
        pos = m.end()
    return tokens


def parse_filter(expr):
    """$filter を条件の木にする: ("or", [..]) / ("and", [..]) / ("cmp", 列, 演算子, 値)"""
    tokens = tokenize(expr)
    pos = [0]

    def peek():
        return tokens[pos[0]][0] if pos[0] < len(tokens) else None

    def take(kind):
        if peek() != kind:
            raise ODataError(f"$filter の書き方が正しくありません: {expr}")
        tok = tokens[pos[0]]
        pos[0] += 1
        return tok[1]

    def parse_or():
        terms = [parse_and()]
        while peek() == "or":
            take("or")
            terms.append(parse_and())
        return terms[0] if len(terms) == 1 else ("or", terms)

    def parse_and():
        terms = [parse_atom()]
        while peek() == "and":
            take("and")
            terms.append(parse_atom())
        return terms[0] if len(terms) == 1 else ("and", terms)

    def parse_atom():
        if peek() == "(":
            take("(")
            node = parse_or()
            take(")")
            return node
        field = take("field")
        op = peek()
        if op not in ("eq", "ne", "gt", "ge", "lt", "le"):
            raise ODataError(f"未対応の演算子です: {expr}")
        take(op)
        return ("cmp", field, op, take("lit"))

    node = parse_or()
    if pos[0] != len(tokens):
        raise ODataError(f"$filter の書き方が正しくありません: {expr}")
    return node


def get_field(item, field):
    if field.lower() == "id":
        return item["Id"]
    return item.get(field)


def compare(actual, op, expected):
    if op == "eq":
        return actual == expected
    if op == "ne":
        return actual != expected
    if actual is None or expected is None:
        return False
    try:
        if isinstance(expected, (int, float)) and not isinstance(expected, bool):
            actual = float(actual)
        return {"gt": actual > expected, "ge": actual >= expected,
                "lt": actual < expected, "le": actual <= expected}[op]
    except (TypeError, ValueError):
        return False


def evaluate(node, item):
    kind = node[0]
    if kind == "or":
        return any(evaluate(n, item) for n in node[1])
    if kind == "and":
        return all(evaluate(n, item) for n in node[1])
    _, field, op, value = node
    return compare(get_field(item, field), op, value)


def apply_query(items, params):
    if "$filter" in params and params["$filter"].strip():
        node = parse_filter(params["$filter"])
        items = [it for it in items if evaluate(node, it)]

    if "$orderby" in params:
        keys = []
        for part in params["$orderby"].split(","):
            tokens = part.split()
            if tokens:
                keys.append((tokens[0], len(tokens) > 1 and tokens[1].lower() == "desc"))

        def sort_key(value):
            if value is None:
                return (0, 0, "")  # SharePoint と同じく null は昇順で先頭
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                return (1, value, "")
            return (2, 0, str(value))

        for field, desc in reversed(keys):  # 安定ソートを後ろのキーから適用
            items.sort(key=lambda it, f=field: sort_key(get_field(it, f)), reverse=desc)

    top = params.get("$top")
    try:
        top = int(top) if top is not None else SP_DEFAULT_TOP
    except ValueError:
        raise ODataError(f"$top が不正です: {top}")
    return items[:max(0, top)]


def apply_select(item, select):
    if not select:
        return dict(item)
    out = {}
    for f in [f.strip() for f in select.split(",") if f.strip()]:
        if f.lower() == "id":
            out["Id"] = item["Id"]
            out["ID"] = item["Id"]
        else:
            out[f] = item.get(f)
    return out


def metadata_level(accept):
    accept = (accept or "").lower()
    if "odata=verbose" in accept:
        return "verbose"
    if "odata=minimalmetadata" in accept or "odata=fullmetadata" in accept:
        return "minimal"
    return "none"


def decorate(item, version, list_title, level):
    if level == "minimal":
        return {"odata.type": f"SP.Data.{list_title}ListItem", "odata.etag": make_etag(version), **item}
    if level == "verbose":
        return {"__metadata": {"etag": make_etag(version), "type": f"SP.Data.{list_title}ListItem"}, **item}
    return item


def etag_matches(if_match, version):
    if if_match is None or if_match.strip() == "*":
        return True
    current = make_etag(version)
    for tag in if_match.split(","):
        tag = tag.strip()
        if tag.startswith("W/"):
            tag = tag[2:]
        if tag == current or tag.strip('"') == str(version):
            return True
    return False


def lan_addresses():
    """同じネットワークの端末から開くための、この PC の IP アドレス（分かる範囲で）"""
    addrs = set()
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("10.255.255.255", 1))  # 実際には送信しない（経路の確認だけ）
            addrs.add(s.getsockname()[0])
    except OSError:
        pass
    return sorted(a for a in addrs if not a.startswith("127."))


# ===================== HTTP ハンドラ =====================

class SharePointMockHandler(http.server.SimpleHTTPRequestHandler):
    LIST_PATTERN = re.compile(r"^.*?/_api/web/lists/getbytitle\('((?:[^']|'')+)'\)/items(?:\((\d+)\))?/?$", re.IGNORECASE)
    APP_PATTERN = re.compile(r"^/SiteAssets/([A-Za-z0-9_-]+)(/.*)?$")
    db_path = DEFAULT_DB
    apps = []
    unique_lists = set()
    port = DEFAULT_PORT
    host = "127.0.0.1"

    def log_message(self, fmt, *args):  # 静的ファイルのログは出さない（API だけ）
        if "/_api/" in (self.path or ""):
            super().log_message(fmt, *args)

    # 開発用なので静的ファイルもキャッシュさせない（別端末で古いJSが残るのを防ぐ）
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    # /SiteAssets/<name>/... を各アプリのフォルダに対応させる
    def translate_path(self, path):
        clean = urllib.parse.unquote(urllib.parse.urlsplit(path).path)
        m = self.APP_PATTERN.match(clean)
        app = next((a for a in self.apps if m and a["name"] == m.group(1)), None)
        if not app:
            return str(SCRIPT_DIR / "__not_found__")
        rel = (m.group(2) or "/").lstrip("/")
        target = (app["root"] / rel).resolve()
        if app["root"] not in target.parents and target != app["root"]:
            return str(SCRIPT_DIR / "__not_found__")  # フォルダの外は見せない
        return str(target)

    def send_sp_json(self, data, status=200, level="none", extra_headers=None):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        odata = {"verbose": "verbose", "minimal": "minimalmetadata"}.get(level, "nometadata")
        self.send_header("Content-Type", f"application/json;odata={odata};charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra_headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def send_sp_error(self, status, message, code="-1, Microsoft.SharePoint.Client.InvalidClientQueryException"):
        self.send_sp_json({"odata.error": {"code": code, "message": {"lang": "ja-JP", "value": message}}}, status)

    def send_no_content(self, headers=None):
        self.send_response(204)
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def send_index(self):
        """アプリの一覧ページ"""
        rows = []
        for a in self.apps:
            ok = (a["root"] / "index.html").exists()
            url = f"/SiteAssets/{a['name']}/index.html"
            link = f'<a href="{html.escape(url)}">{html.escape(a["title"])}</a>' if ok else html.escape(a["title"])
            note = "" if ok else f'<span class="ng">フォルダが見つかりません: {html.escape(str(a["root"]))}</span>'
            rows.append(f"<li>{link} <code>{html.escape(url)}</code> {note}</li>")
        if self.host in ("127.0.0.1", "localhost"):
            lan_html = "<p>この PC からだけ開けます。スマホなどからも開きたい場合は <code>--host 0.0.0.0</code> を付けて起動します（ポートの許可が必要です）。</p>"
        else:
            lan = "".join(f"<li><code>http://{ip}:{self.port}/</code></li>" for ip in lan_addresses()) or "<li>（取得できませんでした）</li>"
            lan_html = f"<p>同じネットワークのスマホなどからは、次のアドレスで開けます：</p><ul>{lan}</ul>"
        body = f"""<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>SharePoint モックサーバー</title>
<style>body{{font-family:sans-serif;max-width:760px;margin:24px auto;padding:0 16px;line-height:1.7;color:#1e293b}}
h1{{font-size:20px}}li{{margin:6px 0}}a{{font-weight:bold;color:#0f766e}}code{{background:#f1f5f9;padding:1px 6px;border-radius:4px;font-size:13px}}
.ng{{color:#b91c1c;font-size:13px}}.note{{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:8px 12px;font-size:14px}}</style>
</head><body><h1>SharePoint モックサーバー</h1>
<p>このサーバーのデータ（SQLite）で動くアプリ：</p><ul>{''.join(rows)}</ul>
<p class="note">URL に <code>?mock=1</code> を付けると、このサーバーを使わず、ブラウザ内のテストデータ（保存されない）で動きます。
画面の左下に「テストデータで動作中」と出ます。</p>
{lan_html}
<p>データを初期化するには、サーバーを止めて <code>python mockserver.py --reset</code> で起動し直してください。</p>
</body></html>""".encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def parse_request_target(self):
        parts = urllib.parse.urlsplit(self.path)
        path = urllib.parse.unquote(parts.path)
        params = {k: v[-1] for k, v in urllib.parse.parse_qs(parts.query, keep_blank_values=True).items()}
        match = self.LIST_PATTERN.match(path)
        if not match:
            return path, params, None, None
        list_title, item_id = match.groups()
        return path, params, list_title.replace("''", "'"), (int(item_id) if item_id else None)

    def read_body(self):
        length = int(self.headers.get("Content-Length", 0) or 0)
        if length <= 0:
            return {}
        body = json.loads(self.rfile.read(length).decode("utf-8"))
        if not isinstance(body, dict):
            raise ValueError("JSON オブジェクトを送信してください")
        for key in ("__metadata", "Id", "ID", "odata.type", "odata.etag"):
            body.pop(key, None)
        return body

    def title_taken(self, conn, list_title, title, exclude_id=None):
        """「一意の値を適用」のまね: 同じ Title の行がほかにあるか"""
        if list_title not in self.unique_lists or title is None:
            return False
        for row_id, data in conn.execute("SELECT id, data FROM sp_items WHERE list_title=?", (list_title,)):
            if row_id != exclude_id and json.loads(data).get("Title") == title:
                return True
        return False

    # ----- GET -----
    def do_GET(self):
        path, params, list_title, item_id = self.parse_request_target()
        if list_title is None:
            if "/_api/" in path:
                return self.send_sp_error(404, f"未対応のAPIです: {path}")
            if path in ("/", "/index.html"):
                return self.send_index()
            m = self.APP_PATTERN.match(path)
            if m:
                rest = m.group(2)
                query = urllib.parse.urlsplit(self.path).query
                if rest is None:  # /SiteAssets/seatmap → /SiteAssets/seatmap/
                    self.send_response(301)
                    self.send_header("Location", path + "/" + ("?" + query if query else ""))
                    self.end_headers()
                    return
            return super().do_GET()

        level = metadata_level(self.headers.get("Accept"))
        with connect(self.db_path) as conn:
            if item_id is not None:
                row = conn.execute("SELECT id, version, data FROM sp_items WHERE list_title=? AND id=?",
                                   (list_title, item_id)).fetchone()
                if not row:
                    return self.send_sp_error(404, f"アイテムが存在しません (Id={item_id})",
                                              "-2130575338, Microsoft.SharePoint.SPException")
                item, version = row_to_item(row)
                out = decorate(apply_select(item, params.get("$select")), version, list_title, level)
                if level == "verbose":
                    out = {"d": out}
                return self.send_sp_json(out, level=level, extra_headers={"ETag": make_etag(version)})

            rows = conn.execute("SELECT id, version, data FROM sp_items WHERE list_title=? ORDER BY id",
                                (list_title,)).fetchall()
        pairs = [row_to_item(r) for r in rows]
        versions = {item["Id"]: version for item, version in pairs}
        try:
            items = apply_query([item for item, _ in pairs], params)
        except ODataError as e:
            return self.send_sp_error(400, str(e))
        out = [decorate(apply_select(it, params.get("$select")), versions[it["Id"]], list_title, level) for it in items]
        if level == "verbose":
            return self.send_sp_json({"d": {"results": out}}, level=level)
        return self.send_sp_json({"value": out}, level=level)

    # ----- POST -----
    def do_POST(self):
        path, params, list_title, item_id = self.parse_request_target()

        if path.rstrip("/").lower().endswith("/_api/contextinfo"):
            return self.send_sp_json({
                "FormDigestValue": DIGEST_VALUE,
                "FormDigestTimeoutSeconds": DIGEST_TIMEOUT_SECONDS,
                "d": {"GetContextWebInformation": {"FormDigestValue": DIGEST_VALUE,
                                                   "FormDigestTimeoutSeconds": DIGEST_TIMEOUT_SECONDS}},
            })

        if list_title is None:
            return self.send_sp_error(400, f"未対応のAPIです: {path}")

        if self.headers.get("X-RequestDigest") != DIGEST_VALUE:
            return self.send_sp_error(403, "X-RequestDigest が無効です。ページを再読み込みしてください。",
                                      "-2130575251, Microsoft.SharePoint.SPException")

        try:
            body = self.read_body()
        except (ValueError, json.JSONDecodeError) as e:
            return self.send_sp_error(400, f"リクエスト本文が不正です: {e}")

        sp_method = (self.headers.get("X-HTTP-Method") or "POST").upper()
        level = metadata_level(self.headers.get("Accept"))
        if_match = self.headers.get("IF-Match")
        duplicate = ("一意の値が適用されている列に、重複した値があります。",
                     "-2130575169, Microsoft.SharePoint.SPDuplicateValuesFoundException")

        with DB_LOCK, connect(self.db_path) as conn:
            if sp_method == "POST" and item_id is None:
                if self.title_taken(conn, list_title, body.get("Title")):
                    return self.send_sp_error(409, *duplicate)
                new_id = conn.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM sp_items WHERE list_title=?",
                                      (list_title,)).fetchone()[0]
                conn.execute("INSERT INTO sp_items (list_title, id, version, data) VALUES (?, ?, 1, ?)",
                             (list_title, new_id, json.dumps(body, ensure_ascii=False)))
                out = decorate({"Id": new_id, "ID": new_id, **body}, 1, list_title, level)
                if level == "verbose":
                    out = {"d": out}
                return self.send_sp_json(out, status=201, level=level, extra_headers={"ETag": make_etag(1)})

            if item_id is None or sp_method not in ("MERGE", "PATCH", "PUT", "DELETE"):
                return self.send_sp_error(400, f"未対応の操作です: {sp_method} {path}")

            row = conn.execute("SELECT version, data FROM sp_items WHERE list_title=? AND id=?",
                               (list_title, item_id)).fetchone()
            if not row:
                return self.send_sp_error(404, f"アイテムが存在しません (Id={item_id})",
                                          "-2130575338, Microsoft.SharePoint.SPException")
            version, data = row
            if not etag_matches(if_match, version):
                return self.send_sp_error(
                    412, "このアイテムは他のユーザーによって変更されています。最新の状態を取得してください。",
                    "-1, Microsoft.SharePoint.Client.ClientServiceException")

            if sp_method == "DELETE":
                conn.execute("DELETE FROM sp_items WHERE list_title=? AND id=?", (list_title, item_id))
                return self.send_no_content()

            if "Title" in body and self.title_taken(conn, list_title, body.get("Title"), exclude_id=item_id):
                return self.send_sp_error(409, *duplicate)
            current = {} if sp_method == "PUT" else json.loads(data)
            current.update(body)
            new_version = version + 1
            conn.execute("UPDATE sp_items SET data=?, version=? WHERE list_title=? AND id=?",
                         (json.dumps(current, ensure_ascii=False), new_version, list_title, item_id))
            return self.send_no_content({"ETag": make_etag(new_version)})


def main():
    parser = argparse.ArgumentParser(description="SharePoint REST API モックサーバー（複数アプリ対応）")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--host", default="127.0.0.1",
                        help="待ち受けるアドレス（既定 127.0.0.1 = この PC だけ。0.0.0.0 でほかの端末からも開ける）")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="配信するアプリの設定（apps.json）")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="データを保存する SQLite ファイル")
    parser.add_argument("--reset", action="store_true", help="DB を初期化してサンプルデータを入れ直す")
    args = parser.parse_args()

    cfg = load_config(args.config)
    init_db(args.db, cfg["apps"], reset=args.reset)

    SharePointMockHandler.db_path = args.db
    SharePointMockHandler.apps = cfg["apps"]
    SharePointMockHandler.unique_lists = cfg["unique"]
    SharePointMockHandler.port = args.port
    SharePointMockHandler.host = args.host
    handler = functools.partial(SharePointMockHandler, directory=str(SCRIPT_DIR))
    server = http.server.ThreadingHTTPServer((args.host, args.port), handler)

    print(f"DB: {args.db.resolve()}")
    print(f"SharePoint モックサーバーを起動しました。ブラウザで開いてください: http://localhost:{args.port}/")
    for a in cfg["apps"]:
        ok = "" if (a["root"] / "index.html").exists() else "  ※ フォルダが見つかりません（apps.json の root を確認）"
        print(f"  {a['title']}: http://localhost:{args.port}/SiteAssets/{a['name']}/index.html{ok}")
    if args.host not in ("127.0.0.1", "localhost"):
        for ip in lan_addresses():
            print(f"  同じネットワークの端末から: http://{ip}:{args.port}/")
    print("止めるときは Ctrl + C")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
