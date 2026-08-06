# mayumi-app

まゆみ助産院のお客様アプリと管理者アプリの公開用リポジトリです。

## 公開内容

- お客様アプリ: `index.html` / `app.js` / `style.css` / `sw.js`
- 管理者アプリ: `admin/`（`admin/index.html` と関連アセット一式）
- お客様向け操作マニュアル: `manuals/`

## 主な公開URL

- お客様アプリ: `/`
- 管理者アプリ: `/admin/`
- 操作マニュアルPDF: `/manuals/お客様アプリ操作マニュアル.pdf`

## 反映方針

- GitHub Pages で `main` ブランチの最新内容を公開します。
- お客様アプリの静的ファイルはリポジトリ直下に配置しています。
- 管理者アプリは `admin/` 配下で完結しており、専用の Service Worker (`admin/admin-sw.js`) と
  マニフェスト (`admin/admin-manifest.json`) を持ちます。スコープが `/admin/` に分かれているため、
  お客様アプリの Service Worker とは競合しません。
- 管理者アプリはもともと別リポジトリ `mayumi-josanin/mayumi-admin-app` にあり、
  2026-08-07 に履歴ごと `admin/` へ統合しました。旧URL `/mayumi-admin-app/` は新URLへリダイレクトします。

## 備考

- バックエンドの Google Apps Script 本体はこの公開リポジトリには含めていません。
- iOS 配布用のネイティブプロジェクト一式も含めていません。
