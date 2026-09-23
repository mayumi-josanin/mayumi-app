#!/usr/bin/env bash
#
# 管理画面と窓口を、いまの develop に合わせる（サーバーPCで流す）。
#
#     bash ~/projects/mayumi-app/server/scripts/deploy.sh
#
# まとめた場所（mayumi）に移したあとは、置き場が変わるだけで中身は同じ:
#     bash ~/projects/mayumi/services/manage/scripts/deploy.sh
#
# なぜ1つにまとめたか（2026-09-22）:
#   コードは箱に直接載せている（docker-compose.yml の `.:/app`）ので、
#   **ビルドのときに集めた CSS と JS は、載せた瞬間に隠れる。**
#   つまり `git pull` だけでは、画面に配られる CSS と JS は古いままになる。
#   実際に、摘要が枠に収まる直しがこれで届かず、iPhone に古い画面が出ていた。
#   手順を覚えていなくても済むように、要るものを全部ここに置く。
#
#   **並びには意味がある。**集め直す（collectstatic）→ 起動し直す、の順。
#   逆にすると、箱は古い一覧を覚えたまま動き続ける（WhiteNoise は起動時に数える）。
#
# お客様の記録には触らない。途中で止まったら、そこで何が起きたかを出して終わる。
#
set -euo pipefail

BRANCH="${1:-develop}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"   # サーバーの置き場（server / services/manage）
ROOT="$(git -C "$HERE" rev-parse --show-toplevel)"        # リポジトリの根っこ
REL="${HERE#"$ROOT"/}"                                    # 根っこから見た、サーバーの置き場

# Git Bash は "/app/x" を "C:/Program Files/Git/app/x" に書き換える。止める。
export MSYS_NO_PATHCONV=1
export MSYS2_ARG_CONV_EXCL="*"

say() { printf '\n== %s\n' "$1"; }

cd "$ROOT"
if [ ! -d "$HERE/apps/manage" ]; then
  echo "ここは管理画面の置き場ではないようです: $HERE" >&2
  exit 1
fi

say "いまの置き場を確かめる"
NOW="$(git rev-parse --abbrev-ref HEAD)"
echo "いま開いている枝: $NOW"
if [ "$NOW" != "$BRANCH" ]; then
  echo "**止めます。**この置き場は $NOW を開いています（$BRANCH を取り込む手順です）。" >&2
  echo "わざと別の枝にしているなら、枝の名前を付けて流してください: bash scripts/deploy.sh $NOW" >&2
  exit 1
fi

say "いまの版を控える"
BEFORE="$(git rev-parse HEAD)"
git --no-pager log --oneline -1

say "$BRANCH を取り込む"
git pull --ff-only "origin" "$BRANCH"
AFTER="$(git rev-parse HEAD)"
git --no-pager log --oneline -1

if [ "$BEFORE" = "$AFTER" ]; then
  echo "（新しいものはありませんでした。念のため、配るものは集め直します）"
fi

cd "$HERE"

# 依存（Pillow など）や箱の作り方が変わったときだけ作り直す。ふだんは数分を使わない
if ! git -C "$ROOT" diff --quiet "$BEFORE" "$AFTER" -- "$REL/Dockerfile" "$REL/pyproject.toml"; then
  say "箱を作り直す（Dockerfile か依存が変わったため。数分かかります）"
  docker compose build web manage cron
fi

say "箱を起こす"
docker compose up -d

say "表の形を合わせる"
docker compose exec -T manage python manage.py migrate --noinput

say "配るもの（CSS・JS・画像）を集め直す"
docker compose exec -T manage python manage.py collectstatic --noinput

say "起動し直す"
docker compose restart manage web

say "配られる中身が、いまのコードと同じか確かめる"
# collectstatic をもう一度「書かずに」流す。0 件なら、集め終わっている印
LEFT="$(docker compose exec -T manage python manage.py collectstatic --noinput --dry-run 2>/dev/null | tail -3 || true)"
printf '%s\n' "$LEFT"
if printf '%s' "$LEFT" | grep -qE '^0 '; then
  echo "→ 配るものは、いまのコードと同じです"
else
  echo "→ **まだ集め切れていないものがあります。**上の行を控えて、開発側に見せてください" >&2
fi

say "出来上がり"
echo "管理画面: https://desktop-rmsk0vg.tail8efe0d.ts.net:10002/manage/"
echo "画面が古いままなら、端末で一度読み込み直してください（iPhone は引き下げて更新）。"
