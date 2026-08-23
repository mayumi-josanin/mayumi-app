"""「知ったきっかけアンケート」まわりの、会員ごとの印を取り込む。

GAS では**表ではなく、会員ごとのスクリプトプロパティ**に入っている。

    SURVEY_ANSWERED:<会員ID>        … 回答した事実
    SURVEY_STAMP:<会員ID>           … お礼スタンプを付けた記録
    SURVEY_STAMP_PENDING:<会員ID>   … 満杯で付けられず保留にした記録
    REWARD_ADMIN_SET:<会員ID>       … 管理者がスタンプを編集した時刻

**移し忘れると、お礼スタンプが二重に付く。**
付与の記録が空に見えるので「まだ付けていない」と判断してしまう。

使いかた:
    python manage.py 会員ごとの印を取り込む 会員ごとの印.json --下見
    python manage.py 会員ごとの印を取り込む 会員ごとの印.json
"""

import json

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_datetime

from apps.members.models import Member

# JSONの名前 → モデルの項目
移すもの = {
    "surveyAnsweredAt": "survey_answered_at",
    "surveyStampGrantedAt": "survey_stamp_granted_at",
    "surveyStampPendingAt": "survey_stamp_pending_at",
    "rewardAdminSetAt": "reward_admin_set_at",
}


def _日時(値):
    """日時に直す。

    GAS は `formatDateTime_(new Date())` の文字列を入れているが、
    **中身が日時でないこともある**（真偽値の代わりに使われている鍵がある）。
    読めなければ「印はある」ものとして**取り込んだ時刻**を入れる。
    空にすると「印が無い」ことになり、二重付与につながる。
    """
    s = str(値 or "").strip()
    if not s:
        return None
    t = parse_datetime(s)
    if t is None:
        d = parse_date(s[:10])
        t = (timezone.datetime.combine(d, timezone.datetime.min.time())
             if d else timezone.now())
    if timezone.is_naive(t):
        t = timezone.make_aware(t)
    return t


class Command(BaseCommand):
    help = "会員ごとのアンケートの印を取り込む"

    def add_arguments(self, parser):
        parser.add_argument("ファイル")
        parser.add_argument("--下見", action="store_true",
                            help="書き込まずに、何がどう変わるかだけ出す")

    def handle(self, *args, **opts):
        下見 = opts["下見"]
        try:
            with open(opts["ファイル"], encoding="utf-8") as f:
                生 = json.load(f)
        except OSError as e:
            raise CommandError(f"読めませんでした: {e}")

        行 = 生.get("members") if isinstance(生, dict) else 生
        if not isinstance(行, list):
            raise CommandError("members が配列ではありません。")

        変わる = []
        いない = []
        読めなかった = 0
        for r in 行:
            会員ID = str(r.get("memberId") or "").strip()
            if not 会員ID:
                continue
            m = Member.objects.filter(member_id=会員ID).first()
            if not m:
                いない.append(会員ID)
                continue

            差 = {}
            for 名, 項目 in 移すもの.items():
                if 名 not in r:
                    continue
                t = _日時(r.get(名))
                if t is None:
                    continue
                if str(r.get(名) or "").strip() and parse_datetime(str(r.get(名)).strip()) is None:
                    読めなかった += 1
                if getattr(m, 項目) != t:
                    差[項目] = t
            if 差:
                変わる.append((m, 差))

        self.stdout.write(f"■ ファイルの中身: {len(行)}名ぶん")
        self.stdout.write(f"  会員が見つからない: {len(いない)}名")
        for i in いない[:10]:
            self.stdout.write(f"    {i}")
        if 読めなかった:
            self.stdout.write(f"  日時として読めなかった値: {読めなかった}件"
                              f"（**印はあるものとして、取り込んだ時刻を入れます**）")
        self.stdout.write("")

        self.stdout.write(f"■ 変わる会員: {len(変わる)}名")
        for 名, 項目 in 移すもの.items():
            n = sum(1 for _, d in 変わる if 項目 in d)
            self.stdout.write(f"    {名}: {n}名")
        self.stdout.write("")
        付与済み = Member.objects.exclude(survey_stamp_granted_at=None).count()
        self.stdout.write(f"  取り込み前の「お礼スタンプ付与済み」: {付与済み}名")

        if 下見:
            self.stdout.write(self.style.WARNING("下見なので書き込んでいません。"))
            return

        with transaction.atomic():
            for m, 差 in 変わる:
                for k, v in 差.items():
                    setattr(m, k, v)
                m.save(update_fields=list(差.keys()) + ["changed_at"])

        self.stdout.write(self.style.SUCCESS(f"■ {len(変わる)}名を更新しました。"))
        self.stdout.write(
            f"  取り込み後の「お礼スタンプ付与済み」: "
            f"{Member.objects.exclude(survey_stamp_granted_at=None).count()}名")
