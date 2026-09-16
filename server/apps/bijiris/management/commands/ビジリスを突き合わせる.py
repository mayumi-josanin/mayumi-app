"""ビジリスの回答の JSON（スプレッドシートの写し）と、サーバーの回答を突き合わせる。**読むだけ。**

    python manage.py ビジリスを突き合わせる ビジリス回答.json

件数と、回答ごとの全項目（取り込みが写す項目すべて）を比べ、違いを1件ずつ出す。
切り替えの前に「サーバーにあるものがスプレッドシートと同じか」を確かめるため。
サーバーが正になったあとは、シート側が古いだけなので違いが出て当たり前（そのときは使わない）。
"""

from django.core.management.base import BaseCommand

from apps.bijiris.models import Response, ResponsePhoto

from ._common import ファイルを開く, 文字
from .ビジリスの回答を取り込む import 写真を集める, 行にする


class Command(BaseCommand):
    help = "ビジリスの回答の JSON とサーバーの回答を突き合わせる（読むだけ）"

    def add_arguments(self, parser):
        parser.add_argument("json_path")

    def handle(self, *args, **options):
        生 = ファイルを開く(options["json_path"])
        一覧 = 生.get("responses") if isinstance(生, dict) else 生
        一覧 = [r for r in (一覧 if isinstance(一覧, list) else []) if isinstance(r, dict) and 文字(r.get("id"))]

        # sheet_row と survey（FK）は比べない。行は取り込みの印、FK はサーバー側の結び
        比べない = {"sheet_row", "survey"}
        シートだけ, 違う, 同じ = [], [], 0
        サーバーのID = set(Response.objects.values_list("response_id", flat=True))
        for r in 一覧:
            rid = 文字(r.get("id"))[:128]
            サーバーのID.discard(rid)
            既存 = Response.objects.filter(response_id=rid).first()
            if not 既存:
                シートだけ.append(rid)
                continue
            値 = 行にする(r)
            差 = [k for k, v in 値.items() if k not in 比べない and getattr(既存, k) != v]
            写真の数 = len(写真を集める(値["answers"], 値["files"]))
            写真の差 = ResponsePhoto.objects.filter(response=既存).count() != 写真の数
            if 差 or 写真の差:
                違う.append((rid, 差 + (["写真の枚数"] if 写真の差 else [])))
            else:
                同じ += 1

        self.stdout.write("")
        self.stdout.write(f"■ ビジリスの回答: JSON {len(一覧)}件 / サーバー {Response.objects.count()}件")
        self.stdout.write(f"    同じ: {同じ}件")
        self.stdout.write(f"    違う: {len(違う)}件")
        for rid, 差 in 違う[:50]:
            self.stdout.write(f"      - {rid}: {'・'.join(差)}")
        self.stdout.write(f"    JSON にだけある: {len(シートだけ)}件 {シートだけ[:10]}")
        self.stdout.write(f"    サーバーにだけある: {len(サーバーのID)}件 {sorted(サーバーのID)[:10]}")
        self.stdout.write("")
        self.stdout.write("  ※ 読むだけです。何も書いていません。")
