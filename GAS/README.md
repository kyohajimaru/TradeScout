# GAS + Googleスプレッドシート版 海外商品リサーチ管理ツール

PCが起動していなくても、Google Apps Script の時間主導トリガーで毎日 JST 09:05 に海外商品候補を5件まで収集し、Googleスプレッドシートへ直接追記します。

## ファイル

- `Code.gs`: Apps Script 本体
- `appsscript.json`: タイムゾーンと権限スコープ設定

## 初期設定

1. Googleスプレッドシートを新規作成します。
2. メニューから `拡張機能` -> `Apps Script` を開きます。
3. `Code.gs` の内容を Apps Script エディタへ貼り付けます。
4. プロジェクト設定で `appsscript.json` を表示し、このフォルダの `appsscript.json` の内容に置き換えます。
5. Apps Script の `プロジェクトの設定` -> `スクリプト プロパティ` に以下を追加します。

| プロパティ | 値 |
| --- | --- |
| `SERPAPI_API_KEY` | SerpAPI のAPIキー |

APIキーはコードに直接書かないでください。

## 初回実行

Apps Script エディタ上で以下の順に実行します。

1. `setupSheets()`
2. `createDailyTrigger()`

初回は権限承認が必要です。

## 手動実行

すぐ収集したい場合は、Apps Script エディタから以下のどちらかを実行します。

- `collectDailyProducts()`
- `manualCollectDailyProducts()`

## 作成されるシート

- `商品候補一覧`
- `カテゴリ設定`
- `収集ログ`
- `除外リスト`

`商品候補一覧` の既存データは削除せず、末尾に追記します。重複チェックは以下の組み合わせで行います。

- 商品ページURL
- 公式URL
- ブランド名 + 商品名

## 運用メモ

- 1回の収集件数は5件です。
- SerpAPI呼び出し回数を抑えるため、初期カテゴリでは1カテゴリ1件ずつ収集します。
- 価格が取れない商品は `ステータス` が `価格確認待ち` になります。
- 利益計算は `商品候補一覧` の数式で自動計算されます。
- `カテゴリ設定` の `有効` を `FALSE` にすると、そのカテゴリは収集対象外になります。
- `除外リスト` にキーワードを追加すると、検索結果の除外判定に使われます。

## トリガー管理

- `createDailyTrigger()`: 毎日 JST 09:05 の時間主導トリガーを作成します。
- `deleteTriggers()`: `collectDailyProducts` の既存トリガーを削除します。

Apps Script の時間主導トリガーは `nearMinute(5)` を使うため、Google側の仕様で数分前後する場合があります。
