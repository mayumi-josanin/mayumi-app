"""スプレッドシートから書き出した JSON を DB に取り込む。

    python manage.py 計測記録を取り込む 測定履歴.json
    python manage.py 計測記録を取り込む 測定履歴.json --下見

「測定ID」で突き合わせるので、**何度実行しても二重に増えない**。
すでにある行は、中身が変わっていれば更新する。

会員IDは、まゆみ側の会員台帳が移るまで空のままでよい。
JSONに memberId が入っていれば、そのまま持つ。
"""

import json
from datetime import datetime
from decimal import Decimal, InvalidOperation

from django.core.management.base import BaseCommand, CommandError
from django.utils import timezone

from apps.measurements.models import Measurement


def 数字にする(value):
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def 日付にする(value):
    if not value:
        return None
    text = str(value)[:10]
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def 日時にする(value):
    if not value:
        return None
    text = str(value).replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        return None
    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt)
    return dt


class Command(BaseCommand):
    help = "スプレッドシートから書き出した計測記録のJSONを取り込む"

    def add_arguments(self, parser):
        parser.add_argument("ファイル", help="測定履歴のJSON")
        parser.add_argument(
            "--下見", action="store_true", dest="dry_run",
            help="書き込まずに、何件入って何件変わるかだけを出す",
        )

    def handle(self, *args, **options):
        path = options["ファイル"]
        下見 = options["dry_run"]

        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except OSError as e:
            raise CommandError(f"読めませんでした: {e}") from e

        rows = data.get("measurements") if isinstance(data, dict) else data
        if not isinstance(rows, list):
            raise CommandError("JSONの形が違います（measurements の配列が要ります）")

        新規 = 0
        更新 = 0
        変わらず = 0
        とばした = []

        for row in rows:
            mid = str(row.get("measurementId") or "").strip()
            name = str(row.get("customerName") or "").strip()
            measured = 日付にする(row.get("measuredOn"))

            # 測定IDと測定日が無いものは、あとで見分けられないので取り込まない。
            if not mid or not measured:
                とばした.append(f"{name or '(名前なし)'} … 測定ID または 測定日 がありません")
                continue

            値 = {
                "member_id": str(row.get("memberId") or "").strip(),
                "customer_name": name,
                "member_number": str(row.get("memberNumber") or "").strip(),
                "measured_on": measured,
                "waist": 数字にする(row.get("waist")),
                "hip": 数字にする(row.get("hip")),
                "thigh_right": 数字にする(row.get("thighRight")),
                "thigh_left": 数字にする(row.get("thighLeft")),
                "whr": 数字にする(row.get("whr")),
                "staff_memo": str(row.get("staffMemo") or ""),
                "created_at": 日時にする(row.get("createdAt")),
                "updated_at": 日時にする(row.get("updatedAt")),
            }

            既存 = Measurement.objects.filter(measurement_id=mid).first()
            if 既存 is None:
                新規 += 1
                if not 下見:
                    Measurement.objects.create(measurement_id=mid, **値)
                continue

            変更あり = any(getattr(既存, k) != v for k, v in 値.items())
            if 変更あり:
                更新 += 1
                if not 下見:
                    for k, v in 値.items():
                        setattr(既存, k, v)
                    既存.save()
            else:
                変わらず += 1

        見出し = "【下見】書き込んでいません" if 下見 else "取り込みました"
        self.stdout.write(f"■ {見出し}")
        self.stdout.write(f"  読んだ行 : {len(rows)}")
        self.stdout.write(f"  新しく入る : {新規}")
        self.stdout.write(f"  中身が変わる : {更新}")
        self.stdout.write(f"  変わらない : {変わらず}")
        if とばした:
            self.stdout.write(f"  とばした : {len(とばした)}")
            for s in とばした:
                self.stdout.write(f"    - {s}")
        if not 下見:
            self.stdout.write(f"  DBの件数 : {Measurement.objects.count()}")
