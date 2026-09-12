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

  # ---- 何と比べるか ----
  #
  # **本番と比べてはいけない。**控えを取ったあとに増えた分まで
  # 「不一致」になる。実際、日中に流したら content_news が 105 対 104 に
  # なり、**正常な控えを「戻せません」と言った**（2026-09-12）。
  # 増えた1件は、その日の 6:51 に投稿されたものだった。
  # 毎回そう出る道具は、誰も信じなくなる。
  #
  # 控えの隣に、取った時点の件数を書き残すようにした（控えを取る.ps1）。
  # あればそれと比べる。**これが本当の「戻せるか」の試験。**
  # 無ければ（古い控え）本番と比べるが、**差は「あとの増減」として扱う。**
  $件数ファイル = [System.IO.Path]::ChangeExtension($控え, ".counts.txt")
  $取った時の件数 = $null
  if (Test-Path $件数ファイル) {
    $取った時の件数 = @{}
    foreach ($行 in Get-Content $件数ファイル) {
      if ("$行" -match '^([A-Za-z0-9_]+)\|(-?\d+)$') { $取った時の件数[$Matches[1]] = [int]$Matches[2] }
    }
    Write-Host "  （控えを取った時点の件数と比べます）"
  } else {
    Write-Host "  （この控えには取った時点の件数がありません。本番と比べ、差は増減として見ます）"
  }
  Write-Host ""

  $名前 = ($本番の表.Keys + $戻した表.Keys) | Sort-Object -Unique |
    Where-Object { $数えない -notcontains $_ }
  $合う = 0
  $欠け = @()      # **表そのものが無い。これは戻せていない**
  $ずれ = @()      # 件数が違う。増減かもしれない

  foreach ($n in $名前) {
    $b = 実数 $試す入れ物 $n
    if ($取った時の件数) {
      $a = if ($取った時の件数.ContainsKey($n)) { $取った時の件数[$n] } else { -1 }
      $見出し = "取った時"
    } else {
      $a = 実数 "mayumi" $n
      $見出し = "本番    "
    }

    if ($b -lt 0) {
      $印 = "欠け"; $欠け += $n
    } elseif ($a -eq $b) {
      $印 = "OK  "; $合う++
    } else {
      $印 = "差   "; $ずれ += ("{0}({1}→{2})" -f $n, $a, $b)
    }
    Write-Host ("  {0} {1,-32} {2} {3,6}  戻した {4,6}" -f $印, $n, $見出し, $a, $b)
  }

  Write-Host ""
  if ($欠け.Count -gt 0) {
    Write-Host "■ **戻せていない表があります: $($欠け -join '・')**"
    Write-Host "   この控えは使えません。原因を調べてください。"
  } elseif ($ずれ.Count -eq 0) {
    Write-Host "■ 全 $($名前.Count) 表とも件数が一致しました。**この控えは戻せます。**"
    Write-Host "  （$($数えない -join '・') は毎分変わる一時的な表なので数えていません）"
  } elseif ($取った時の件数) {
    # 取った時点と比べて違うなら、**それは本当に戻せていない。**
    Write-Host "■ **取った時点と件数が違います: $($ずれ -join '・')**"
    Write-Host "   この控えでは元に戻し切れません。原因を調べてください。"
  } else {
    # 本番と比べただけなので、差は「控えを取ったあとの増減」で説明がつく。
    Write-Host "■ 全 $($名前.Count) 表が戻りました。**この控えは戻せます。**"
    Write-Host "   本番と件数が違う表: $($ずれ -join '・')"
    Write-Host "   これは控えを取ったあとの増減です（異常ではありません）。"
    Write-Host "   確かめたいときは、控えを取り直してから試してください。"
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
