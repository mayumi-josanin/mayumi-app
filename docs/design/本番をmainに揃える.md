# 本番を main に揃える（9/19 のあと、すぐ）

**2026-09-17。院長のご希望は「9/19 以降、できる段階で早急に」。**

いまは逆になっている。お客様アプリは `main` から公開されているのに、
**サーバーの管理画面と窓口は `develop` で動いている。** 作りかけが本番で動く形。
予約システムのように「develop で作る → 確かめる → main で公開」に揃える。

## いまの並び（2026-09-17 時点）

| もの | 動いているもの | 出どころ |
|---|---|---|
| お客様アプリ | `main` | GitHub Pages（`main` の `/`） |
| 管理画面（サーバー :10002） | **`develop`** | `~/projects/mayumi-app` が develop |
| 窓口（サーバー :10001 の API） | **`develop`** | 同上 |
| 予約システム | `main` | `~/projects/mayumi-reserve` |

`develop` は `main` より **169件先**。`main` は `develop` の先祖なので、
**そのまま進められる**（作り直しや突き合わせは要らない）。

## 出すと何が変わるか

- **管理画面**: この夏に作ったものが全部（会員管理・商品・NEWS・お知らせ・カレンダー・
  通知・システム管理・公式サイト管理・開発管理・ビジリス管理・出勤と給料）。
  すでに本番で使っているものなので、**見た目は変わらない**。
- **お客様アプリ**: `index.html` と `sw.js` の版の印だけ（`app.js?v=20260913-2` / `mayumi-app-v74`）。
  中身の直しは、カレンダーの区分・ブログのリンクを外す・ホーム画面のアイコン。
  **どれも `main` に別の形ですでに入っている**ので、お客様の画面は実質変わらない。
- **はじめ方の案内**: `はじめ方A4.html` が `アプリQR.html` に置き換わる（9/13 に作り直したもの）。

## 手順

### ① develop を main に出す（公開）

```
cd ~/Desktop/mayumi-app
git checkout main
git merge --ff-only develop
git push origin main
```

`--ff-only` が通らなければ**止めること**。誰かが main を直に触っている。

出したら、お客様アプリを開いて確かめる:
- https://mayumi-josanin.github.io/mayumi-app/ が開く
- ホーム画面に追加したアプリからも開く（版が上がるので一度読み直しが要る）

### ② サーバーの本番を main に切り替える

```
cd ~/projects/mayumi-app
git fetch origin
git checkout main
git pull --ff-only origin main
bash ~/app_manage_deploy.sh
```

**①のあとでないと巻き戻る。** 順番を守ること。

### ③ 確認用の管理画面を立てる

予約システムで作ったもの（`mayumi-reserve/docs/確認用の予約システム.md`）と同じ形。

- 置き場所: `~/projects/mayumi-app-kakunin`（`develop`）
- 箱の名前: `-p mayumi-kakunin`（本番とぶつからないように）
- 口: web 8762 / manage 8763 / db 5762、住所は `tailscale serve --https=10005`
- **LINE・メール・通知（OneSignal）の合言葉は空にする。** 確認用からお客様に飛ばさない
- データは本番の写し（`pg_dump` で読むだけ）
- 反映の仕掛け: `~/app_kakunin_deploy.sh`

### ④ 以後の流れ

```
develop で作る → 確認用（:10005）で見る → main に入れる → 本番へ出す
```

お客様アプリは `main` に入れた時点で自動的に公開される（GitHub Pages）。
**管理画面と窓口は、`main` に入れたあと、サーバーで反映の一手を打つ**（表を足す作業が混ざるため）。

## 気をつけること

- **9/19 の会員移行が済むまでやらない。** ①はお客様アプリの公開でもある。移行の前に動かさない
- ①の前に、サーバーのデータベースの控えを取る（`~/projects/mayumi-app/server/scripts/` の控え）
- 戻したくなったら、①は `git revert`、②は `git checkout develop` で戻せる

関連: [1つにまとめる](1つにまとめる.md) / [9-19当日の手順](9-19当日の手順.md)
