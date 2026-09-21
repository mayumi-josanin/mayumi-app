"""システム管理。**中身は旧管理アプリ（admin/index.html の #page-system-management）と同じ。**

旧アプリは GAS の getAdminDashboardData（buildAdminDashboardData_）の作り置きを出していた。
ここではサーバーのデータベースから**同じ定義で**その場で数える（作り置きは要らない）。

  集計カード   会員数 / 公開予約 / 重複候補 / Push失敗
  アラート     バックアップ未更新 / 未対応注文 / Push失敗 / 在庫警告 / 重複会員候補 /
               長期間未対応注文 / 未受取特典 / 未公開予約（GAS と同じ順・同じ文言）
  公開予約一覧 NEWS・ショップ・カレンダー・ホーム（メニュー）の公開開始日時が未来のもの

## バックアップだけは GAS と事情が違う

旧アプリの「バックアップ」はスプレッドシートを Google ドライブへ写すもので、
GAS のトリガーとスクリプトプロパティを見ていた。**サーバーからはそれが見えない。**
サーバー側の控えは、院のPCで scripts/控えを取る.ps1 が docker の中の pg_dump で取り、
Google ドライブか server/backups に置いている。ここで分かるのは server/backups
（環境変数 BACKUP_DIR で変えられる）に置かれた控えだけ。**分からないことは分からないと出す。**

「今すぐバックアップ」は、pg_dump がこの入れ物（コンテナ）に入っているときだけ押せる。
python:3.12-slim の像には入っていないので、本番では文言だけになる。
"""

import os
import re
import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_datetime
from django.views.decorators.http import require_POST

from apps.content.models import CalendarEvent, Menu, News, Product, PushNotice
from apps.gasapi import admin_member, admin_product
from apps.records.models import BackupRecord

from .permissions import owner_required
from .views_order import _まとめ as _注文をまとめる
from .views_qrcode import アプリの住所, 住所を決める

# GAS の定数と同じ（管理者・お客様.js 277〜279行）
未対応注文の日数 = 3          # STALE_PENDING_ORDER_DAYS
未受取特典の日数 = 30         # STALE_UNUSED_REWARD_DAYS
未公開予約の時間 = 24         # STALE_UNPUBLISHED_SCHEDULE_HOURS
バックアップが古い時間 = 36   # getBackupStatus の staleThreshold
PUSH失敗 = "送信失敗"         # PUSH_STATUS_FAILED


# ---- 日時の読み方 ----

def _日時を読む(v):
    """GAS の parseLooseDateToTimestamp_ にあたる。読めなければ None。

    特典の獲得日は '2026/04/05 18:25' や '2026-04-05T18:25:26+09:00' など形が揃っていない。
    """
    if not v:
        return None
    if hasattr(v, "tzinfo"):
        return v if timezone.is_aware(v) else timezone.make_aware(v)
    s = str(v).strip().replace("/", "-")
    d = parse_datetime(s)
    if d is None:
        日 = parse_date(s.split(" ")[0].split("T")[0])
        if 日 is None:
            return None
        d = datetime(日.year, 日.month, 日.day)
    return d if timezone.is_aware(d) else timezone.make_aware(d)


# ---- 重複会員候補（GAS の buildDuplicateUsersFromRows_ と同じ）----

def _電話を寄せる(v):
    s = re.sub(r"\D", "", str(v or ""))
    if len(s) >= 9 and not s.startswith("0"):
        s = "0" + s
    return s


def _氏名を寄せる(v):
    return re.sub(r"[\s　]+", "", str(v or "").strip())


def _生年月日を寄せる(v):
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", s)
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""


def _重複の組(会員):
    """電話番号 → 氏名+生年月日 → 氏名 の順で組を作り、同じ顔ぶれは1組に数える。"""
    組 = []
    見た = set()

    def 足す(顔ぶれ):
        if len(顔ぶれ) < 2:
            return
        鍵 = "|".join(sorted(u["memberId"] for u in 顔ぶれ))
        if 鍵 in 見た:
            return
        見た.add(鍵)
        組.append(顔ぶれ)

    電話, 氏名生年月日, 氏名 = {}, {}, {}
    for u in 会員:
        p, n, b = _電話を寄せる(u.get("phone")), _氏名を寄せる(u.get("name")), _生年月日を寄せる(u.get("birthday"))
        if p:
            電話.setdefault(p, []).append(u)
        if n and b:
            氏名生年月日.setdefault(n + "|" + b, []).append(u)
        if n:
            氏名.setdefault(n, []).append(u)
    for 表 in (電話, 氏名生年月日, 氏名):
        for 顔ぶれ in 表.values():
            足す(顔ぶれ)
    return 組


# ---- 公開予約 ----

# (元, 種類の名前, 題名の欄, 題名が空のときの名前) … GAS の getPublishSchedule と同じ並び・同じ名前
公開予約の元 = [
    (News, "NEWS", "title", "NEWS"),
    (Product, "ショップ", "name", "商品"),
    (CalendarEvent, "カレンダー", "title", "イベント"),
    (Menu, "ホーム", "name", "メニュー"),
]


def _公開予約(いま):
    """公開開始日時が未来のもの（消したものは除く）。近い順。"""
    出 = []
    for モデル, 種類, 欄, 既定 in 公開予約の元:
        for r in モデル.objects.filter(deleted=False, publish_at__gt=いま).order_by("publish_at"):
            出.append({"source": 種類, "title": (getattr(r, 欄) or "").strip() or 既定, "publishAt": r.publish_at})
    return sorted(出, key=lambda x: x["publishAt"])


def _未公開予約(いま):
    """予定時刻を24時間以上過ぎたのに非公開のまま（GAS の getOverdueScheduledPublishes_）。"""
    境目 = いま - timedelta(hours=未公開予約の時間)
    出 = []
    for モデル, 種類, 欄, 既定 in 公開予約の元:
        for r in モデル.objects.filter(deleted=False, published=False, publish_at__isnull=False, publish_at__lte=境目):
            出.append({"source": 種類, "title": (getattr(r, 欄) or "").strip() or 種類, "publishAt": r.publish_at})
    return sorted(出, key=lambda x: x["publishAt"])


# ---- 在庫警告（GAS の getProductInventoryAlerts_）----

def _在庫警告():
    出 = []
    for p in admin_product.一覧()["products"]:
        在庫 = p.get("stockQty") or 0
        閾値 = p.get("lowStockThreshold") or 0
        if p.get("isSoldOut") or (閾値 > 0 and 0 < 在庫 <= 閾値):
            出.append(p)
    return 出[:10]


# ---- 未受取特典（GAS の getStaleUnusedRewards_）----

def _未受取特典(会員, いま):
    境目 = いま - timedelta(days=未受取特典の日数)
    出 = []
    for u in 会員:
        for r in u.get("rewards") or []:
            if not isinstance(r, dict) or r.get("used"):
                continue
            獲得 = _日時を読む(r.get("earnedDate"))
            if not 獲得 or 獲得 > 境目:
                continue
            出.append({"memberId": u["memberId"], "name": u.get("name") or "",
                       "rewardName": str(r.get("rewardName") or "特典"), "earnedDate": 獲得})
    return sorted(出, key=lambda x: x["earnedDate"])


# ---- 控え（バックアップ）----

def _控えの置き場() -> Path:
    return Path(os.environ.get("BACKUP_DIR") or (settings.BASE_DIR / "backups"))


def _控えの状況(いま):
    """server/backups にある pg_dump の控えを見る。置き場が無ければ「分からない」。"""
    置き場 = _控えの置き場()
    if not 置き場.is_dir():
        return {"visible": False, "dir": str(置き場), "latest": None, "stale": False, "files": []}
    files = []
    for f in 置き場.glob("mayumi-*.dump"):
        try:
            files.append({"name": f.name, "at": timezone.make_aware(datetime.fromtimestamp(f.stat().st_mtime)),
                          "kb": round(f.stat().st_size / 1024, 1)})
        except OSError:
            continue
    files.sort(key=lambda x: x["at"], reverse=True)
    最終 = files[0]["at"] if files else None
    return {"visible": True, "dir": str(置き場), "latest": 最終,
            "stale": (最終 is None) or (最終 < いま - timedelta(hours=バックアップが古い時間)), "files": files[:10]}


def _接続の設定():
    return settings.DATABASES["default"]


def _pg_dumpの場所():
    return shutil.which("pg_dump")


def _控えを取れるか():
    """この入れ物から pg_dump で控えを取れるか。取れないなら、その理由。"""
    if "postgresql" not in _接続の設定().get("ENGINE", ""):
        return False, "このサーバーは PostgreSQL につながっていないため、サーバーからは控えを取れません。"
    if not _pg_dumpの場所():
        return False, ("サーバーの入れ物（コンテナ）に pg_dump が入っていないため、ここからは控えを取れません。"
                       "控えは院のPCで scripts/控えを取る.ps1 が取ります。")
    return True, ""


# ---- 画面 ----

@owner_required
def system_view(request):
    いま = timezone.now()
    会員 = admin_member.一覧()["users"]
    注文 = _注文をまとめる()
    受付中 = [o for o in 注文 if o["status"] == "受付中"]
    境目 = いま - timedelta(days=未対応注文の日数)
    長期未対応 = sorted([o for o in 受付中 if o["date"] and o["date"] <= 境目], key=lambda o: o["date"])
    未受取 = _未受取特典(会員, いま)
    重複 = _重複の組(会員)
    控え = _控えの状況(いま)
    公開予約 = _公開予約(いま)
    未公開 = _未公開予約(いま)
    push失敗 = PushNotice.objects.filter(deleted=False, status=PUSH失敗).count()
    在庫 = _在庫警告()

    # GAS の buildAdminDashboardData_ と同じ順・同じ文言。
    # 「日次バックアップ（トリガー未設定）」は GAS のトリガーの話なので、サーバーからは出せない。
    alerts = []
    if 控え["visible"] and 控え["stale"]:
        alerts.append(("warning", "バックアップ未更新", f"{バックアップが古い時間}時間以上バックアップが更新されていません"))
    if 受付中:
        alerts.append(("info", "未対応注文", f"{len(受付中)}件の受付中注文があります"))
    if push失敗:
        alerts.append(("warning", "Push失敗", f"{push失敗}件の送信失敗があります"))
    if 在庫:
        alerts.append(("warning", "在庫警告", f"{len(在庫)}件の商品で在庫警告があります"))
    if 重複:
        alerts.append(("warning", "重複会員候補", f"{len(重複)}組の重複候補があります"))
    if 長期未対応:
        alerts.append(("warning", "長期間未対応注文", f"{len(長期未対応)}件の受付中注文が{未対応注文の日数}日以上経過しています"))
    if 未受取:
        alerts.append(("warning", "未受取特典", f"{len(未受取)}件の特典が{未受取特典の日数}日以上未受取です"))
    if 未公開:
        alerts.append(("warning", "未公開予約", f"{len(未公開)}件の公開予約が予定時刻を過ぎても非公開のままです"))

    取れる, 取れない理由 = _控えを取れるか()
    return render(request, "manage/system.html", {
        # お客様アプリの住所（旧管理アプリの「初期設定 > アプリ公開用URL」）。QRコード案内がこれを使う
        "app_url": アプリの住所(),
        "generated_at": いま,
        "summary": [("会員数", len(会員)), ("公開予約", len(公開予約)), ("重複候補", len(重複)), ("Push失敗", push失敗)],
        "alerts": alerts,
        "stale_rewards": 未受取[:10],
        "overdue_publishes": 未公開[:10],
        "upcoming_publishes": 公開予約[:12],
        "backup": 控え,
        # スプレッドシートの控えの記録（GAS の BACKUP_LOG をサーバーへ取り込んだもの）。
        # **取り込んだ時点までしか無い。**いまの状況は旧管理アプリで見る。
        "sheet_backups": list(BackupRecord.objects.order_by("-created_at", "-sheet_row")[:10]),
        "can_backup": 取れる,
        "cannot_backup_reason": 取れない理由,
    })


@require_POST
@owner_required
def system_backup(request):
    """今すぐバックアップ。pg_dump で控えを1つ書くだけ（データベースには何も書かない）。"""
    取れる, 理由 = _控えを取れるか()
    if not 取れる:
        messages.error(request, 理由)
        return redirect("manage:system")
    db = _接続の設定()
    置き場 = _控えの置き場()
    置き場.mkdir(parents=True, exist_ok=True)
    ファイル = 置き場 / f"mayumi-{timezone.localtime().strftime('%Y%m%d-%H%M')}.dump"
    env = dict(os.environ, PGPASSWORD=str(db.get("PASSWORD") or ""))
    cmd = [_pg_dumpの場所(), "-h", str(db.get("HOST") or "localhost"), "-p", str(db.get("PORT") or "5432"),
           "-U", str(db.get("USER") or "postgres"), "-Fc", "-f", str(ファイル), str(db.get("NAME"))]
    try:
        r = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=300)
    except (OSError, subprocess.TimeoutExpired) as e:
        messages.error(request, f"バックアップに失敗しました: {e}")
        return redirect("manage:system")
    # 中身が空なら取れていないのと同じ（控えを取る.ps1 と同じ判定）。気づかず何日も過ぎるのが一番困る。
    if r.returncode != 0 or not ファイル.exists() or ファイル.stat().st_size < 1024:
        if ファイル.exists():
            ファイル.unlink()
        messages.error(request, "バックアップに失敗しました" + (f": {r.stderr.strip()[:200]}" if r.stderr else ""))
        return redirect("manage:system")
    messages.success(request, f"バックアップを作成しました（{ファイル.name}）")
    return redirect("manage:system")


@require_POST
@owner_required
def system_app_url(request):
    """お客様アプリの住所を決める。

    旧管理アプリは、この住所をそのパソコンのブラウザの中（localStorage）に置いていた。
    そのため**パソコンを変えると消え、QRコードが出なくなった**。ここでは設定の置き場
    （records.AppSetting）に入れるので、どの端末から入っても同じ住所になる。
    """
    住所 = str(request.POST.get("app_url") or "").strip()
    if 住所 and not 住所.startswith(("http://", "https://")):
        messages.error(request, "住所は https:// から始まる形で入れてください。")
        return redirect("manage:system")
    住所を決める(住所)
    messages.success(request, "お客様アプリの住所を保存しました。" if 住所 else "お客様アプリの住所を空にしました。")
    return redirect("manage:system")
