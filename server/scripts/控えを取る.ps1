# まゆみ助産院データベースの控えを取る。
#
#   .\控えを取る.ps1                     … 既定の置き場へ取る
#   .\控えを取る.ps1 -置き場 "D:\控え"    … 置き場を指定する
#   .\控えを取る.ps1 -残す日数 30         … 何日分残すか（既定14日）
#
# なぜ要るのか:
#   データベースを自宅のPCに置くため、そのPCが壊れたら中身ごと無くなる。
#   会員情報は作り直せない。だから、別の場所（Googleドライブ）へ毎日写す。
#
# 置き場は、Googleドライブのフォルダが見つかればそこを既定にする。
# 見つからなければPCの中に置くが、それは「同じPCの中」なので、
# PCが壊れたときの備えにはならない。その旨を最後に伝える。

param(
  [string]$置き場 = "",
  [int]$残す日数 = 14
)

$ErrorActionPreference = "Stop"

# このスクリプトの1つ上（server フォルダ）で docker compose を動かす。
$サーバー = Split-Path -Parent $PSScriptRoot
Set-Location $サーバー

# ---- 置き場を決める ----
if (-not $置き場) {
  # Googleドライブ（パソコン版）が入っていれば、その中に置く。
  # 入っていればマイドライブが G:\ などに現れる。
  $候補 = @(
    "$env:USERPROFILE\Google ドライブ\まゆみ助産院\データベース控え",
    "$env:USERPROFILE\GoogleDrive\まゆみ助産院\データベース控え",
    "G:\マイドライブ\まゆみ助産院\データベース控え",
    "G:\My Drive\まゆみ助産院\データベース控え"
  )
  foreach ($c in $候補) {
    $親 = Split-Path -Parent (Split-Path -Parent $c)
    if (Test-Path $親) { $置き場 = $c; break }
  }
}
if (-not $置き場) {
  $置き場 = Join-Path $サーバー "backups"
  $ドライブ外 = $true
} else {
  $ドライブ外 = $false
}

if (-not (Test-Path $置き場)) {
  New-Item -ItemType Directory -Path $置き場 -Force | Out-Null
}

# ---- 控えを取る ----
$日時 = Get-Date -Format "yyyyMMdd-HHmm"
$ファイル = Join-Path $置き場 "mayumi-$日時.dump"

Write-Host "■ 控えを取ります"
Write-Host "  置き場: $置き場"

# PowerShell の > は文字コードを変えてしまい、中身が壊れる。
# cmd のリダイレクトを使って、バイトをそのまま書き出す。
$cmd = "docker compose exec -T db pg_dump -U postgres -Fc mayumi > `"$ファイル`""
cmd /c $cmd

if (-not (Test-Path $ファイル)) {
  throw "控えのファイルができませんでした。docker compose ps でデータベースが動いているか確かめてください。"
}

$大きさ = (Get-Item $ファイル).Length
if ($大きさ -lt 1024) {
  # 中身が空なら、取れていないのと同じ。気づかず何日も過ぎるのが一番困る。
  Remove-Item $ファイル -Force
  throw "控えの中身が空でした（$大きさ バイト）。取れていません。"
}

Write-Host "  できました: $(Split-Path -Leaf $ファイル)（$([math]::Round($大きさ/1KB,1)) KB）"

# ---- 古いものを片付ける ----
$境目 = (Get-Date).AddDays(-$残す日数)
$消した = 0
Get-ChildItem -Path $置き場 -Filter "mayumi-*.dump" | Where-Object { $_.LastWriteTime -lt $境目 } | ForEach-Object {
  Remove-Item $_.FullName -Force
  $消した++
}
if ($消した -gt 0) { Write-Host "  $残す日数 日より古い控えを $消した 件片付けました" }

$残り = (Get-ChildItem -Path $置き場 -Filter "mayumi-*.dump").Count
Write-Host "  いま残っている控え: $残り 件"

if ($ドライブ外) {
  Write-Host ""
  Write-Host "※ Googleドライブのフォルダが見つからなかったので、PCの中に置きました。"
  Write-Host "   このままだと、PCが壊れたときに控えも一緒に無くなります。"
  Write-Host "   Googleドライブ（パソコン版）を入れるか、-置き場 で別の場所を指定してください。"
}
