"""仕入値のJSONを取り込む。

    python manage.py 仕入を取り込む 仕入.json --下見
    python manage.py 仕入を取り込む 仕入.json

**シートの行番号で突き合わせるので、何度実行しても二重に増えない。**

仕入値は**商品名で商品マスタと結びついている**（見出しが「商品名（完全一致）」）。
IDに直したくなるが、直すと名前が少し違うだけの商品が迷子になる。
名前のまま持ち、**結びついたかどうかを記録して、結びつかないものは
必ず名前を挙げて報告する。**黙って通すと、あとで粗利を出したときに
「なぜかこの商品だけ原価が出ない」と探すことになる。
"""

import json
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from apps.content.models import Product
from apps.records.models import SupplierPrice


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値):
    if 値 is None or 値 == "":
        return None
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


def 金額(値):
    """仕入値。**整数に丸めない。**

    売上の単価で、まとめ買いの割り算（3個で6995円）が入っていた前例がある。
    仕入値はいま全件が整数だが、同じことが起きても桁を落とさないようにする。
    """
    if 値 is None or 値 == "":
        return None
    try:
        d = Decimal(str(値))
    except (InvalidOperation, TypeError, ValueError):
        return None
    return d.quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)


def 円(n):
    if n is None:
        return "—"
    return f"{int(n):,}円" if n == int(n) else f"{n}円"


class Command(BaseCommand):
    help = "仕入値のJSONを取り込む"

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

        行 = 生.get("supplier_prices") or []
        下見 = options["preview"]

        # 商品マスタの名前を先に集める。削除済みも含めて照合する
        # （消えた商品の仕入値も、記録としては残っているため）。
        商品 = {p.name.strip(): p for p in Product.objects.all()}

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        採用, つながらない = [], []
        名前の数 = {}

        for r in 行:
            row = 数(r.get("row"))
            名 = 文字(r.get("product_name"))
            if not row or not 名:
                飛ばした += 1
                continue
            採用.append(r)
            名前の数[名] = 名前の数.get(名, 0) + 1

            相手 = 商品.get(名)
            if 相手 is None:
                つながらない.append((row, 名))

            値 = {
                "product_name": 名[:255],
                "price": 金額(r.get("price")),
                "memo": 文字(r.get("memo")),
                "matched_product": 相手,
            }

            既存 = SupplierPrice.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 仕入値: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号か商品名が空）: {飛ばした}件")

        空 = sum(1 for r in 採用 if 金額(r.get("price")) is None)
        ゼロ = sum(1 for r in 採用 if 金額(r.get("price")) == 0)
        端数 = sum(1 for r in 採用
                   if (v := 金額(r.get("price"))) is not None and v != int(v))
        self.stdout.write(f"    仕入値: 空欄 {空}件 / 0 {ゼロ}件"
                          + (f" / **割り切れない値 {端数}件**" if 端数 else ""))

        # 同じ商品名が2行あると、どちらの仕入値か決まらない。
        重なり = {k: v for k, v in 名前の数.items() if v > 1}
        if 重なり:
            self.stdout.write("")
            self.stdout.write("    **同じ商品名が複数行あります。どちらの値か決まりません:**")
            for 名, n in sorted(重なり.items()):
                self.stdout.write(f"      {名} … {n}行")

        # 商品マスタと結びついたか。**見つからないものは必ず名前を挙げる。**
        つながった = len(採用) - len(つながらない)
        self.stdout.write("")
        self.stdout.write(f"  ● 商品マスタとの結びつき: {つながった}/{len(採用)}件")
        if つながらない:
            self.stdout.write(f"    **結びつかない {len(つながらない)}件:**")
            for row, 名 in つながらない:
                self.stdout.write(f"      {row}行 「{名}」")
            self.stdout.write("      → 商品マスタに同じ名前がありません。")
            self.stdout.write("        取り扱いをやめた商品か、名前の書き方が違うかです。")
            self.stdout.write("        **仕入値は記録として入りますが、商品には繋がりません。**")

        # 逆向きも見る。原価の分からない商品を先に知っておく。
        名前ら = {文字(r.get("product_name")) for r in 採用}
        原価なし = [p.name for p in Product.objects.filter(deleted=False)
                    if p.name.strip() not in 名前ら]
        if 原価なし:
            self.stdout.write("")
            self.stdout.write(f"  ● 仕入値の無い商品: {len(原価なし)}件")
            for 名 in 原価なし:
                self.stdout.write(f"      「{名}」")
            self.stdout.write("      → この商品は原価が分からないので、粗利が出せません。")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for row, 値 in 新規:
                SupplierPrice.objects.create(sheet_row=row, **値)
            for row, 値 in 更新:
                # ForeignKey は update() に渡せる形に直す。
                値2 = dict(値)
                相手 = 値2.pop("matched_product")
                SupplierPrice.objects.filter(sheet_row=row).update(
                    matched_product=相手, **値2)

        self.stdout.write("")
        self.stdout.write(f"    → いま {SupplierPrice.objects.count()}件")
