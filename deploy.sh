#!/usr/bin/env bash
set -euo pipefail

# ============================================================
# ./deploy.sh "変更メモ"
#   1回のコマンドで フロント(GitHub Pages) と GAS(Apps Script) の
#   両方を反映する。GAS は毎回「同じデプロイID」に上書きするので
#   Web アプリ URL は変わらず、shared/gas-config.js の修正も不要。
# ============================================================

message="${1:-Update survey app}"
ROOT="$(cd "$(dirname "$0")" && pwd)"

# --- 設定（基本さわらない）-------------------------------------------------
# 本番用デプロイID。gas-config.js の /macros/s/●●●/exec の ●●● 部分と同じ。
GAS_DEPLOYMENT_ID="AKfycbyZfGreoBxYYaB5i0NX1Hw_9eAMD5Q1YTfzKORYyLc2-TQt6R6xblkUNyW8SlEih4QW"
# --------------------------------------------------------------------------

# 1) 構文チェック（壊れたまま反映しないように）
if command -v npm >/dev/null 2>&1; then
  echo "▶ 構文チェック (npm run check)..."
  npm run check
  echo "✓ 構文OK"
fi

# 2) GAS を Apps Script に反映（clasp が使えるときだけ）
if command -v clasp >/dev/null 2>&1; then
  echo "▶ GAS を Apps Script に反映 (clasp push + 同一URLへ再デプロイ)..."
  ( cd "$ROOT/gas" \
      && clasp push -f \
      && clasp deploy -i "$GAS_DEPLOYMENT_ID" -d "$message" )
  echo "✓ GAS 反映完了（URLは固定のまま）"
else
  echo "※ clasp 未導入のため GAS 反映をスキップ（フロントのみ反映します）"
  echo "  GAS も変更した場合は clasp を入れて再実行してください。"
fi

# 3) フロントを GitHub Pages に反映
cd "$ROOT"
git add .
if git diff --cached --quiet; then
  echo "GitHub 側にコミットする変更はありません。"
else
  git commit -m "$message"
  git push
  echo "✓ GitHub Pages へ push 完了"
fi

echo ""
echo "公開URL:"
echo "  お客様: https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/"
echo "  管理:   https://mayumi-josanin.github.io/mayumi_bijiris/admin-app/"
echo "（反映に1〜2分。PWAキャッシュがある場合は再読み込み）"
