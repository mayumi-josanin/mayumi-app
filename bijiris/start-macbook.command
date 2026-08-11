#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

echo "Mayumi Bijiris を起動します。"
echo "このウィンドウを閉じるとアプリは停止します。"
echo "お客様には customer-app の URL だけを共有してください。"
echo

npm start
