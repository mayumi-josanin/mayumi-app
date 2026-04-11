# Mayumi Bijiris Survey App

まゆみ助産院で使うアンケート Web アプリです。お客様はスマホからお客様アプリに回答し、管理者は別の管理アプリで回答を確認・管理します。

## ファイル

- `customer-app/`: お客様に共有するアンケート回答アプリ
- `admin-app/`: 管理者が回答を管理するアプリ
- `shared/`: お客様アプリと管理アプリで共通利用する API 接続コード
- `gas/`: Google Apps Script の保存・管理 API
- `default-surveys.js`: 初期アンケート定義
- `index.html`: アプリ選択用の案内ページ
- `server.js`: ローカル確認用の Node.js API と静的ファイル配信
- `data/`: ローカル確認時の回答データ保存場所

## 公開URL

- お客様アプリ: `https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/`
- 管理アプリ: `https://mayumi-josanin.github.io/mayumi_bijiris/admin-app/`
- 回答保存スプレッドシート: `https://docs.google.com/spreadsheets/d/1oDNTqlvKv1rGOGXIpnzPlegpFDeQ0WHGRLuY3ZAnZYc/edit`
- Apps Script プロジェクト: `https://script.google.com/d/13qSqUDPU9WSHMX_LrmU1zFP3f64VdQaXVqUKgRu4uFb8bqnSkieqIMEa/edit`

## Google Sheets + Apps Script 運用

回答保存と管理ログインは Google Apps Script を使います。Render などの有料サーバーは不要です。

Apps Script WebアプリURLは `shared/gas-config.js` に設定します。

```js
window.MAYUMI_GAS_WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbyZfGreoBxYYaB5i0NX1Hw_9eAMD5Q1YTfzKORYyLc2-TQt6R6xblkUNyW8SlEih4QW/exec";
```

Apps Script 側では次を行います。

1. Apps Script プロジェクトを開きます。
2. `setup` 関数を選び、実行して Google Sheets と Google Drive へのアクセスを承認します。
3. プロジェクトの設定から「スクリプト プロパティ」を開き、管理者情報を設定します。
4. 「デプロイ」から Web アプリを作成または更新します。
5. 実行するユーザーは「自分」、アクセスできるユーザーは「全員」にします。
6. Web アプリURLが変わった場合は `shared/gas-config.js` を更新し、`./deploy.sh "Update Apps Script URL"` で GitHub Pages へ反映します。

固定の `SPREADSHEET_ID` にアクセスできない場合は、Apps Script 実行アカウント側の Drive に新しい保存用スプレッドシートを自動作成して継続します。

推奨するスクリプト プロパティ:

- `SPREADSHEET_ID`: `1oDNTqlvKv1rGOGXIpnzPlegpFDeQ0WHGRLuY3ZAnZYc`
- `ADMIN_USERNAME`: 管理アプリのログインID
- `ADMIN_PASSWORD`: 管理アプリのパスワード
- `TOKEN_SECRET`: 管理ログイン用の長いランダム文字列

未設定の場合の初期管理者ログイン:

- ID: `mayumi2026`
- PASS: `3939`

公開運用では `ADMIN_USERNAME` と `ADMIN_PASSWORD` を必ず変更してください。

## 保存内容

スプレッドシートには次のシートを用意しています。

- `回答一覧`: 全アンケート回答の一覧、管理状況、回答JSON、写真JSON
- `ビジリス施術アンケート`: 施術アンケート別の回答
- `ビジリス回数券終了アンケート`: 回数券終了アンケート別の回答
- `モニター終了アンケート`: モニター終了アンケート別の回答

写真質問はスマホから画像を選択できます。送信時に Apps Script が Google Drive に写真ファイルを保存し、回答には Drive のURLを記録します。保存先は Google Drive の `bijiris` フォルダ配下で、お客様ごとの名前フォルダに分けます。

## アンケート

お客様アプリには次の3種類を表示します。

- ビジリス施術アンケート
- ビジリス回数券終了アンケート
- モニター終了アンケート

質問項目は `default-surveys.js` と `gas/Code.gs` の両方に定義しています。質問を変更する場合は両方を揃えて更新してください。

## スマホに追加

お客様用 URL をスマホのブラウザで開きます。

- iPhone: 共有ボタンから「ホーム画面に追加」
- Android: ブラウザの案内またはメニューから「アプリをインストール」

お客様にはお客様アプリのURLだけを共有してください。管理アプリのURLは管理者のみが利用します。

## ローカル確認

Node.js 18 以上で確認できます。

```bash
npm start
```

ブラウザで次を開きます。

- `http://localhost:3000/customer-app/`
- `http://localhost:3000/admin-app/`

ローカルの Node.js API だけで確認する場合、回答データは `data/db.json` に保存されます。GitHub Pages公開時は Google Apps Script を使います。

## デプロイ

GitHub Pages へ反映する場合は、次を使います。

```bash
./deploy.sh "変更内容のメモ"
```

今後ファイルを編集したら、このコマンドで GitHub へ push し、公開ページに反映します。
