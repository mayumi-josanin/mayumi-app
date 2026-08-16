# 控えが本当に戻せるかを、本番に触れずに確かめる。
#
#   .\控えを試す.ps1                  … いちばん新しい控えで試す
#   .\控えを試す.ps1 -控え "（path）"  … 控えを指定する
#
# **本番のデータベースには一切書きません。**
# 別名の入れ物（mayumi_restore_test）を作ってそこへ戻し、中身を数えて、
# 最後にその入れ物ごと捨てます。
#
# なぜ要るのか:
#   控えは「取れている」だけでは足りない。戻せない控えは控えではない。
#   本番へ戻して確かめるのは怖くてできないので、使い捨ての入れ物で試す。
#   会員データを入れる前に、一度は通しておく。

param(
  [string]$控え = ""
)

$ErrorActionPreference = "Stop"

$サーバー = Split-Path -Parent $PSScriptRoot
Set-Location $サーバー

$試す入れ物 = "mayumi_restore_test"

# ---- 控えを選ぶ ----
if (-not $控え) {
  $候補 = Get-ChildItem (Join-Path $サーバー "backups") -Filter "mayumi-*.dump" -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending | Select-Object -First 1
  if (-not $候補) { throw "控えが見つかりません。先に 控えを取る.ps1 を実行してください。" }
  $控え = $候補.FullName
}
if (-not (Test-Path $控え)) { throw "控えが見つかりません: $控え" }

$情報 = Get-Item $控え
Write-Host "■ 試す控え: $($情報.Name)"
Write-Host "  取った日時: $($情報.LastWriteTime)"
Write-Host "  大きさ: $([math]::Round($情報.Length/1KB,1)) KB"
Write-Host ""

# ---- 使い捨ての入れ物を作る ----
Write-Host "■ 使い捨ての入れ物を用意します（$試す入れ物）"
docker compose exec -T db psql -U postgres -c "DROP DATABASE IF EXISTS $試す入れ物;" | Out-Null
docker compose exec -T db psql -U postgres -c "CREATE DATABASE $試す入れ物;" | Out-Null

try {
  # ---- 控えを入れ物へ送り込む ----
  # PowerShell の > は文字コードを変えてしまうので、cmd のリダイレクトを使う。
  $cmd = "docker compose exec -T db sh -c ""cat > /tmp/restore-test.dump"" < ""$控え"""
  cmd /c $cmd

  Write-Host "■ 戻しています…"
  # pg_restore は所有者や権限まわりで警告を出すことがあるが、
  # 中身が戻っているかは次の件数で判断する。
  cmd /c "docker compose exec -T db pg_restore -U postgres -d $試す入れ物 /tmp/restore-test.dump 2>&1" | Out-Null

  # ---- 本番と件数を突き合わせる ----
  Write-Host ""
  Write-Host "■ 本番と戻したものを、表ごとに数えて比べます"
  Write-Host ""

  $表を数える = @"
SELECT relname, n_live_tup FROM pg_stat_user_tables ORDER BY relname;
"@

  function 件数($db) {
    $出力 = docker compose exec -T db psql -U postgres -d $db -t -A -F "|" -c $表を数える
    $表 = @{}
    foreach ($行 in $出力) {
      $t = "$行".Trim()
      if ($t -match '^([A-Za-z0-9_]+)\|(-?\d+)$') { $表[$Matches[1]] = [int]$Matches[2] }
    }
    return $表
  }

  # pg_stat_user_tables は概算なので、実数で数え直す。
  function 実数($db, $表名) {
    $出力 = docker compose exec -T db psql -U postgres -d $db -t -A -c "SELECT count(*) FROM `"$表名`";"
    $t = ("$出力" -join "").Trim()
    if ($t -match '^\d+$') { return [int]$t }
    return -1
  }

  $本番の表 = 件数 "mayumi"
  $戻した表 = 件数 $試す入れ物

  # 数えない表。
  # django_cache は呼び出し回数を数える一時的な入れ物で、毎分中身が変わる。
  # 控えを取った時点と今とで違って当然なので、一致を求めると必ず失敗する。
  $数えない = @("django_cache", "django_session")

  $名前 = ($本番の表.Keys + $戻した表.Keys) | Sort-Object -Unique |
    Where-Object { $数えない -notcontains $_ }
  $合う = 0; $合わない = @()

  foreach ($n in $名前) {
    $a = 実数 "mayumi" $n
    $b = 実数 $試す入れ物 $n
    $印 = if ($a -eq $b) { "OK " } else { "NG " }
    if ($a -eq $b) { $合う++ } else { $合わない += $n }
    Write-Host ("  {0} {1,-34} 本番 {2,6}  戻した {3,6}" -f $印, $n, $a, $b)
  }

  Write-Host ""
  if ($合わない.Count -eq 0) {
    Write-Host "■ 全 $($名前.Count) 表とも件数が一致しました。**この控えは戻せます。**"
    Write-Host "  （$($数えない -join '・') は毎分変わる一時的な表なので数えていません）"
  } else {
    Write-Host "■ **一致しない表があります: $($合わない -join '・')**"
    Write-Host "   この控えでは元に戻し切れません。原因を調べてください。"
  }
}
finally {
  # 何があっても使い捨ての入れ物は捨てる。残すと次に紛らわしい。
  Write-Host ""
  Write-Host "■ 使い捨ての入れ物を片付けます"
  docker compose exec -T db psql -U postgres -c "DROP DATABASE IF EXISTS $試す入れ物;" | Out-Null
  docker compose exec -T db rm -f /tmp/restore-test.dump 2>&1 | Out-Null
  Write-Host "  片付けました（本番のデータベースには何も書いていません）"
}
