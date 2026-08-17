"""売上のJSONを取り込む。

    python manage.py 売上を取り込む 売上.json --下見
    python manage.py 売上を取り込む 売上.json

**種別とシートの行番号の組で突き合わせるので、何度実行しても二重に増えない。**

行番号だけでは足りない。MENU_REVENUE の2行目と PRODUCT_REVENUE の2行目が
どちらも「2」になり、あとから来たほうが先のを上書きしてしまう。
"""

import json
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.records.models import RevenueRecord


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値):
    """空欄は None のまま返す。**0 と空欄を混ぜない。**

    原価0（仕入れの無い教室など）を「記録していない」に丸めると、
    粗利が変わってしまう。
    """
    if 値 is None or 値 == "":
        return None
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


def 日付(値):
    s = 文字(値)[:10]
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def 日時(値):
    s = 文字(値)
    if not s:
        return None
    try:
        d = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if timezone.is_naive(d):
        d = timezone.make_aware(d, timezone.get_current_timezone())
    return d


def 円(n):
    return "—" if n is None else f"{n:,}円"


class Command(BaseCommand):
    help = "売上のJSONを取り込む"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview",
                            help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        try:
            with open(options["json_path"], encoding="utf-8") as f:
                生 = json.load(f)
        except (OSError, ValueError) as e:
            raise CommandError(f"読み込めませんでした: {e}")

        行 = 生.get("revenues") or []
        下見 = options["preview"]

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        採用 = []

        for r in 行:
            row = 数(r.get("row"))
            種別 = 文字(r.get("kind"))
            名 = 文字(r.get("name"))
            日 = 日付(r.get("recorded_on"))

            if not row or 種別 not in (RevenueRecord.MENU, RevenueRecord.PRODUCT):
                飛ばした += 1
                continue
            # 名前も日付も無い行は、書きかけの空行。
            if not 名 and not 日:
                飛ばした += 1
                continue
            採用.append(r)

            値 = {
                "recorded_on": 日,
                "name": 名[:255],
                "quantity": 数(r.get("quantity")),
                "unit_price": 数(r.get("unit_price")),
                "unit_cost": 数(r.get("unit_cost")),
                "memo": 文字(r.get("memo")),
                "deleted": bool(文字(r.get("deleted"))),
                "deleted_at": 日時(r.get("deleted_at")),
            }

            既存 = RevenueRecord.objects.filter(kind=種別, sheet_row=row).first()
            if not 既存:
                新規.append((種別, row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((種別, row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 売上: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（種別・行番号が無い、名前も日付も空）: {飛ばした}件")

        for 種別 in (RevenueRecord.MENU, RevenueRecord.PRODUCT):
            分 = [r for r in 採用 if 文字(r.get("kind")) == 種別]
            if not 分:
                continue
            self.stdout.write("")
            self.stdout.write(f"  ● {種別}: {len(分)}件")
            for 鍵, 名 in (("quantity", "数"), ("unit_price", "単価"), ("unit_cost", "原価")):
                空 = sum(1 for r in 分 if 数(r.get(鍵)) is None)
                ゼロ = sum(1 for r in 分 if 数(r.get(鍵)) == 0)
                self.stdout.write(f"      {名}: 空欄 {空}件 / 0 {ゼロ}件")
            # 金額は、数と単価がそろっている行だけで出す。
            出せる = [r for r in 分
                      if 数(r.get("quantity")) is not None and 数(r.get("unit_price")) is not None]
            売上 = sum(数(r["quantity"]) * 数(r["unit_price"]) for r in 出せる)
            self.stdout.write(f"      売上の合計: {円(売上)}"
                              + (f"（数か単価が空で出せない行 {len(分) - len(出せる)}件）"
                                 if len(出せる) != len(分) else ""))

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for 種別, row, 値 in 新規:
                RevenueRecord.objects.create(kind=種別, sheet_row=row, **値)
            for 種別, row, 値 in 更新:
                RevenueRecord.objects.filter(kind=種別, sheet_row=row).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {RevenueRecord.objects.count()}件")
