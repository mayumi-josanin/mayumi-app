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

# ---- Googleドライブへ送る ----
#
# PCの中に置くだけでは、そのPCが壊れたときに控えも一緒に無くなる。
# 会員情報は作り直せないので、置き場をPCの外に持つ。
#
# GAS に送る向きにしてある。逆（GASが取りに来る）にすると、サーバー側に
# 「データベース全体を返す窓口」を作ることになり、合鍵ひとつで会員情報が
# 丸ごと持ち出せる口をインターネットに開けてしまう。
#
# .env の BACKUP_UPLOAD_KEY と MAYUMI_GAS_URL が要る。無ければ送らずに知らせる。

$env設定 = @{}
Get-Content (Join-Path $サーバー ".env") | ForEach-Object {
  if ($_ -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$') { $env設定[$Matches[1]] = $Matches[2].Trim() }
}
$送り先 = $env設定["MAYUMI_GAS_URL"]
$送信鍵 = $env設定["BACKUP_UPLOAD_KEY"]

if (-not $送り先 -or -not $送信鍵) {
  Write-Host ""
  Write-Host "※ Googleドライブへ送る設定がありません（.env の MAYUMI_GAS_URL / BACKUP_UPLOAD_KEY）。"
  Write-Host "   このままだと、PCが壊れたときに控えも一緒に無くなります。"
} else {
  try {
    $本文 = @{
      type     = "uploadBackup"
      key      = $送信鍵
      filename = (Split-Path -Leaf $ファイル)
      content  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ファイル))
    } | ConvertTo-Json -Compress

    # GAS は必ず転送（302）を返す。PowerShell は転送のときに POST を GET に
    # 変えてしまうため、そのままだと doGet の応答を受け取ってしまう。
    # doGet も status: ok を返すので、**送れていないのに成功に見える。**
    # （実際にそれで「送りました」と表示しながら0件だった）
    # 転送先を自分で受け取り、そこへ POST し直す。
    # PowerShell 5.1 は -MaximumRedirection 0 のとき、302 を「例外」として投げる。
    # 転送先は、その例外が持つ応答のヘッダから取り出す。
    $転送先 = $null
    try {
      # -UseBasicParsing は必須。付けないと古いIEの部品を使おうとして、
      # この環境では NullReferenceException になる（応答すら返らない）。
      $一次 = Invoke-WebRequest -Uri $送り先 -Method Post -Body $本文 -UseBasicParsing `
        -ContentType "application/json" -TimeoutSec 180 -MaximumRedirection 0
      $転送先 = $一次.Headers["Location"]
    } catch {
      $res = $_.Exception.Response
      if ($res) { $転送先 = $res.Headers["Location"] }
    }
    if (-not $転送先) { $転送先 = $送り先 }

    $生 = Invoke-WebRequest -Uri $転送先 -Method Post -Body $本文 -UseBasicParsing `
      -ContentType "application/json" -TimeoutSec 180
    $応答 = $生.Content | ConvertFrom-Json

    # 「送ったファイル名が返ってきたか」まで確かめる。
    # status だけ見ると、別の処理の応答でも成功に見えてしまう。
    if ($応答.status -eq "ok" -and $応答.saved -eq (Split-Path -Leaf $ファイル)) {
      Write-Host "  ドライブへ送りました: $($応答.saved)（$([math]::Round($応答.bytes/1KB,1)) KB）"
      if ($応答.removed -gt 0) { Write-Host "  ドライブの古い控えを $($応答.removed) 件片付けました" }
    } else {
      $理由 = if ($応答.message) { $応答.message } else { "応答が想定と違います（送れていません）" }
      Write-Host "  **ドライブへ送れませんでした: $理由**"
    }
  } catch {
    # 送れなくてもPCの中の控えは残る。ここで止めない。
    Write-Host "  **ドライブへ送れませんでした: $($_.Exception.Message)**"
  }
}

if ($ドライブ外) {
  Write-Host ""
  Write-Host "※ PCの中の置き場は $置き場 です（ドライブへ送れていれば、こちらは予備）。"
}
