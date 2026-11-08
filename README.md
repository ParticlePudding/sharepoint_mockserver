# SharePoint モックサーバー

SharePoint Online の REST API（リストの読み書き）をまねる、**開発用の小さなサーバー**です。
SharePoint がなくても、[座席表（動静表）](../動静表/動静表Webアプリ) と [会議室予約](../会議室予約) を、**本物と同じ通信のしかた**で動かせます。

- Python 3 の標準機能だけで動きます（追加のインストールは不要）。
- データは同じフォルダの SQLite ファイル（`mock_sharepoint.db`）に保存されます。サーバーを止めても残り、複数のブラウザ・端末で共有されます。
- 本番では使いません（SharePoint に置くのは各アプリのファイルだけです）。

---

## 起動する

```bash
cd SharePointモックサーバー
python mockserver.py
```

ブラウザで <http://localhost:8000/> を開くと、アプリの一覧が出ます。

| アプリ | URL |
|---|---|
| 座席表（動静表） | <http://localhost:8000/SiteAssets/seatmap/index.html> |
| 会議室予約 | <http://localhost:8000/SiteAssets/roombook/index.html> |

止めるときは、起動した画面で `Ctrl` + `C` を押します。

- 既定では **この PC（127.0.0.1）からの接続だけ**を受け付けます。ほかの端末に向けてポートを開けません。
- URL に `?mock=1` を付けると、アプリはこのサーバーを使わず、ブラウザ内のテストデータ（保存されない）で動きます。そのときは画面の左下に「**テストデータで動作中（保存されません）**」と出ます。

### スマホなど別の端末から開く（任意）
`--host 0.0.0.0` を付けて起動すると、同じネットワークの端末から、起動時に表示される `http://<PCのIPアドレス>:8000/` で開けます。
ポートを外から受け付けることになるので、会社のルールで許可されている場合だけ使ってください。

```bash
python mockserver.py --host 0.0.0.0
```

---

## データ

- 初めて起動したとき、各アプリのサンプルデータを入れます。
  - 座席表：`js/sample-data.js` の社員108名と人数表のグループ
  - 会議室予約：起動した日を基準にした架空の予約（繰り返し予約を含む）
- **最初の状態に戻す**：サーバーを止めて、`--reset` を付けて起動します。

```bash
python mockserver.py --reset
```

| オプション | 内容 |
|---|---|
| `--port 8080` | ポート番号を変える（既定 8000） |
| `--host 0.0.0.0` | ほかの端末からの接続も受け付ける（既定 `127.0.0.1` = この PC だけ） |
| `--reset` | データを消してサンプルを入れ直す |
| `--db ファイル名` | データを保存するファイルを変える |
| `--config ファイル名` | アプリの設定ファイルを変える（既定 `apps.json`） |

---

## 配信するアプリの設定（`apps.json`）

```json
{
  "apps": [
    { "name": "seatmap",  "title": "座席表（動静表）", "root": "../動静表/動静表Webアプリ", "seed": "seatmap" },
    { "name": "roombook", "title": "会議室予約",       "root": "../会議室予約",            "seed": "roombook" }
  ],
  "uniqueTitleLists": ["会議室予約", "AppSettings"]
}
```

| キー | 内容 |
|---|---|
| `name` | URL の `/SiteAssets/<name>/` の部分（英数字・`-`・`_`） |
| `title` | 一覧ページに出す名前 |
| `root` | アプリのフォルダ（`index.html` のある場所）。**このファイルからの相対パス** |
| `seed` | 初回に入れるサンプルデータの種類（`seatmap` / `roombook`。不要なら書かない） |
| `uniqueTitleLists` | `Title` 列を「一意の値を適用する」にしたリスト。重複した `Title` の追加を 409 で断る |

各アプリをこれとは別の場所に置いた場合（GitHub から別々に取得した場合など）は、`root` を書き換えてください。

---

## まねしている SharePoint の機能

| 機能 | 内容 |
|---|---|
| `GET /_api/web/lists/getbytitle('リスト名')/items` | `$filter`（`eq` `ne` `gt` `ge` `lt` `le`、`and` `or`、括弧）・`$orderby`・`$top`・`$select` |
| `GET …/items(ID)` | 1件の取得（`ETag` ヘッダー付き） |
| `POST /_api/contextinfo` | FormDigest の発行 |
| `POST …/items` | 追加（201） |
| `POST …/items(ID)` + `X-HTTP-Method: MERGE` / `DELETE` | 更新・削除。`IF-Match` が今の ETag と違えば **412**（楽観ロック） |
| `Accept: …odata=minimalmetadata` | 各行に `odata.etag` を付ける |

- リストは、最初に書き込んだ時点で自動的にできます（作る手順は不要）。
- URL の形（`/SiteAssets/<フォルダ>/index.html`）も本物と同じなので、各アプリの「サイトURLの自動判定」もそのまま動きます。
- 権限・認証・リストの列の種類・しきい値（5,000行）はまねしていません。本番前に、本物の SharePoint でも確認してください。

---

## フォルダ構成

```
├ mockserver.py     サーバー本体
├ apps.json         配信するアプリの設定
├ README.md
└ mock_sharepoint.db  データ（自動で作られる。Git には含めない）
```
