# 控えからデータベースを戻す。
#
#   .\控えから戻す.ps1 -控え "G:\マイドライブ\...\mayumi-20260816-0300.dump"
#   .\控えから戻す.ps1 -控え "...\mayumi-....dump" -本当に戻す
#
# **既定では中身を見るだけで、書き戻しません。**
# 戻すと、いまのデータベースの中身は控えの時点まで巻き戻ります。
# 取り違えて実行すると、その間の記録が消えるため、二段構えにしてある。
#
# 控えは、取れているだけでは足りない。戻せることを一度は確かめておくこと。
# 「戻せない控え」は控えではない。

param(
  [Parameter(Mandatory=$true)][string]$控え,
  [switch]$本当に戻す
)

$ErrorActionPreference = "Stop"

$サーバー = Split-Path -Parent $PSScriptRoot
Set-Location $サーバー

if (-not (Test-Path $控え)) { throw "控えが見つかりません: $控え" }

$情報 = Get-Item $控え
Write-Host "■ 控え: $($情報.Name)"
Write-Host "  取った日時: $($情報.LastWriteTime)"
Write-Host "  大きさ: $([math]::Round($情報.Length/1KB,1)) KB"
Write-Host ""

if (-not $本当に戻す) {
  Write-Host "  中身を見るだけにしています（何も書き戻していません）。"
  Write-Host ""
  Write-Host "  含まれている表:"
  $一時 = Join-Path $env:TEMP "mayumi-restore-check.dump"
  Copy-Item $控え $一時 -Force
  cmd /c "docker compose exec -T db sh -c ""cat > /tmp/check.dump"" < ""$一時"""
  cmd /c "docker compose exec -T db pg_restore -l /tmp/check.dump" | Select-String "TABLE DATA" | ForEach-Object {
    Write-Host "    $($_.Line.Trim())"
  }
  Remove-Item $一時 -Force -ErrorAction SilentlyContinue
  Write-Host ""
  Write-Host "  本当に戻すときは -本当に戻す を付けてください。"
  Write-Host "  **いまの中身は、この控えの時点まで巻き戻ります。**"
  exit 0
}

# 戻す前に、いまの中身の控えを取る。戻し先を間違えたときの最後の逃げ道。
Write-Host "■ 念のため、いまの中身の控えを先に取ります"
& (Join-Path $PSScriptRoot "控えを取る.ps1")
Write-Host ""

Write-Host "■ 書き戻します"
$一時 = Join-Path $env:TEMP "mayumi-restore.dump"
Copy-Item $控え $一時 -Force
cmd /c "docker compose exec -T db sh -c ""cat > /tmp/restore.dump"" < ""$一時"""
cmd /c "docker compose exec -T db pg_restore -U postgres -d mayumi --clean --if-exists /tmp/restore.dump"
Remove-Item $一時 -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "■ 戻しました。アプリから見て、記録が揃っているか確かめてください。"
