# 移転用の転送ページ

まゆみ助産院アプリのリポジトリへビジリスを移したあと、
**旧URLを開いた方を新URLへ案内する**ためのファイル。

## 使いどき

`mayumi-app` の `main` を公開し、新URLで動作を確認できてから。
**それより前に設置すると、まだ動かない場所へ案内してしまう。**

## 設置の手順

    cp _移転用/customer-app.html customer-app/index.html
    cp _移転用/admin-app.html    admin-app/index.html
    git add customer-app/index.html admin-app/index.html
    git commit -m "旧URLを新URLへ案内する"
    git push origin main

`deploy.sh` は使わないこと。GASまで反映され、未反映の作業も一緒に出てしまう。

## 転送先

- お客様 https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/
  → https://mayumi-josanin.github.io/mayumi-app/bijiris/customer-app/
- 管理   https://mayumi-josanin.github.io/mayumi_bijiris/admin-app/
  → https://mayumi-josanin.github.io/mayumi-app/bijiris/admin-app/

## 注意

- 6秒後に自動で移る。押していただくボタンも置いてある。
- ホーム画面に追加済みのアイコンは旧URLを指したまま。追加し直しが要る。
- 記録はドメインが同じなので引き継がれる。
- しばらく（数週間）様子を見て、誰も旧URLを使っていないことを確かめてから
  リポジトリの削除を検討すること。削除すると転送も効かなくなる。
