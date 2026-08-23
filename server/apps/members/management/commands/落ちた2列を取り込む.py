"""会員データの **メモ** と **通知の届け先** を取り込む。

最初の「会員を取り込む」で、この2列が落ちていた（2026-08-23に判明）。

  ・メモ … そもそも書き出していなかった
  ・通知の届け先 … 書き出してはいたが、取り込み側が `真偽()` に通していた。
    購読IDは "true" でも "1" でもないので **偽に落ちていた。**

**通知の届け先は、配信の宛先そのもの。**失ったまま切り替えると、
お知らせが誰にも届かなくなる。

使いかた:
    python manage.py 落ちた2列を取り込む 落ちた2列.json --下見
    python manage.py 落ちた2列を取り込む 落ちた2列.json
"""

import json

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from apps.members.models import Member


def _文字(値):
    if 値 is None:
        return ""
    return str(値).strip()


def _通知が入か(届け先):
    """届け先の中身から、通知オンかどうかを決める。

    GAS は `!!row[USER_COL.PUSH - 1]` と書いている（7933行）。
    セルから読むと真偽値の false がそのまま来るので、GAS では偽になる。

    **こちらは文字で受け取る。**"false" という文字列は、そのままでは真になって
    しまうので、明示的に偽へ寄せる。書き出し側でも空にしているが、
    **片方だけに頼らない。**
    """
    s = 届け先.strip()
    if not s:
        return False
    if s.lower() in ("false", "0", "none", "null", "undefined"):
        return False
    return True


class Command(BaseCommand):
    help = "会員データのメモと通知の届け先を取り込む"

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
        for r in 行:
            会員ID = _文字(r.get("memberId"))
            if not 会員ID:
                continue
            m = Member.objects.filter(member_id=会員ID).first()
            if not m:
                いない.append(会員ID)
                continue

            メモ = _文字(r.get("memo"))
            届け先 = _文字(r.get("pushSubscription"))
            入 = _通知が入か(届け先)

            差 = {}
            if m.memo != メモ:
                差["memo"] = メモ
            if m.push_subscription != 届け先:
                差["push_subscription"] = 届け先
            if m.push_enabled != 入:
                差["push_enabled"] = 入
            if 差:
                変わる.append((m, 差))

        self.stdout.write(f"■ ファイルの中身: {len(行)}件")
        self.stdout.write(f"  会員が見つからない: {len(いない)}件")
        for i in いない[:10]:
            self.stdout.write(f"    {i}")
        self.stdout.write("")

        メモが付く = sum(1 for _, d in 変わる if d.get("memo"))
        届け先が付く = sum(1 for _, d in 変わる if d.get("push_subscription"))
        通知が入る = sum(1 for _, d in 変わる if d.get("push_enabled") is True)
        通知が切れる = sum(1 for _, d in 変わる if d.get("push_enabled") is False)

        self.stdout.write(f"■ 変わる会員: {len(変わる)}名")
        self.stdout.write(f"    メモが入る:      {メモが付く}名")
        self.stdout.write(f"    届け先が入る:    {届け先が付く}名")
        self.stdout.write(f"    通知オンになる:  {通知が入る}名")
        self.stdout.write(f"    通知オフになる:  {通知が切れる}名")
        self.stdout.write("")
        self.stdout.write(f"  取り込み前の通知オン: "
                          f"{Member.objects.filter(push_enabled=True).count()}名")

        if 下見:
            self.stdout.write(self.style.WARNING("下見なので書き込んでいません。"))
            return

        with transaction.atomic():
            for m, 差 in 変わる:
                for k, v in 差.items():
                    setattr(m, k, v)
                m.save(update_fields=list(差.keys()) + ["changed_at"])

        self.stdout.write(self.style.SUCCESS(f"■ {len(変わる)}名を更新しました。"))
        self.stdout.write(f"  取り込み後の通知オン: "
                          f"{Member.objects.filter(push_enabled=True).count()}名")
