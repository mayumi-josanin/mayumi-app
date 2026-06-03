# CLAUDE.md — mayumi_bijiris 開発ガイド

このリポジトリで作業するときは、必ずこのファイルのルールに従ってください。

## プロジェクト概要
まゆみ助産院の「ビジリス」アンケート Web アプリ。お客様用と管理者用の2つの PWA があり、
データ保存・管理ログイン・写真保存は Google Apps Script + Google スプレッドシート + Google Drive で動く。
- リポジトリ: https://github.com/mayumi-josanin/mayumi_bijiris （branch: main）
- 公開URL（お客様）: https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/
- 公開URL（管理）:   https://mayumi-josanin.github.io/mayumi_bijiris/admin-app/

## ディレクトリ構成と役割
- customer-app/ … お客様用 PWA（index.html / app.js / styles.css / sw.js / manifest / icons / push）
- admin-app/    … 管理者用 PWA（index.html / app.js / styles.css / sw.js / manifest / icons）
- shared/api.js … 両アプリ共通の API 接続層（GAS or ローカルを自動判定）
- shared/gas-config.js … GAS Web App URL と OneSignal App ID（URL固定運用なので基本さわらない）
- gas/Code.gs   … 本番バックエンド（Apps Script）。回答保存・認証・写真・お客様管理など全API
- default-surveys.js … 初期アンケート定義
- server.js / data/ … ローカル確認用の代替バックエンド（本番では未使用）
- scripts/ … 補助スクリプト　deploy.sh … フロント+GAS をまとめて反映するスクリプト

## デプロイ（重要）
反映は `./deploy.sh "変更メモ"` の1コマンドで完結する。中で次を自動実行する:
1. npm run check で構文確認
2. GAS を Apps Script に反映: gas/ で `clasp push -f` → `clasp deploy -i <固定デプロイID>` で同一URLへ再デプロイ
3. フロントを GitHub Pages に反映: git add/commit/push
- 固定デプロイIDに上書きするので Web アプリ URL は変わらず、shared/gas-config.js の修正は不要。
- 一度だけの準備: `npm install -g @google/clasp` と `clasp login`（Apps Script API を有効化済みであること）。
- clasp が未導入の環境ではフロントのみ反映され、GAS は反映されない点に注意。

## 編集時の必須ルール
- 機能追加はフロント（app.js 等）と GAS（Code.gs）の両方の修正が必要になることが多い。
  どちらを触る必要があるか、編集前に必ず明示すること。
- アンケートの質問定義は default-surveys.js と gas/Code.gs の2か所にある。変更時は必ず両方を揃える。
- PWA キャッシュ対策: 配信ファイルを変えたら sw.js のキャッシュ名（バージョン）を上げる。
- 秘密情報をコードにコミットしない。管理ID/パスワード等は Apps Script のスクリプトプロパティ
  （ADMIN_USERNAME / ADMIN_PASSWORD / TOKEN_SECRET）で管理する。.gitignore は .env と data/*.json を除外済み。
- お客様情報はフルネームのみ（メールアドレス欄なし）。仕様を勝手に増やさない。

## 作業フロー（毎回これに沿う）
1. やることを宣言: customer-app / admin-app / shared / gas のどこを、どう変えるか。
2. 編集する。
3. `npm run check` で構文確認。
4. `./deploy.sh "変更メモ"` で反映（フロント+GASを一括）。
5. 公開URL（customer / admin）を開いて動作確認。GAS変更時は保存・取得も確認。

## やってはいけないこと
- gas/Code.gs を変えただけで「反映済み」と言わない（deploy.sh の再デプロイ手順が必要）。
- 既存の他機能のフィールドや保存スキーマを断りなく変更・削除しない。
- 過去の正しい実装を、根拠なく作り直さない。
