from decimal import ROUND_HALF_UP, Decimal

from django.db import models


class RevenueRecord(models.Model):
    """売上の記録。MENU_REVENUE 112行 と PRODUCT_REVENUE 65行 を1つにまとめたもの。

    **2枚を1表にする。**月ごとの売上を出すときに、毎回2つを足し合わせる
    必要がなくなる。院長が見たいのは「今月いくら入ったか」であって、
    それがメニューか商品かは内訳の話。

    列の名前が2枚で違うので、寄せるときに取り違えないこと（2026-08-17 の下見）:

    | ここでの名前 | MENU_REVENUE | PRODUCT_REVENUE |
    |---|---|---|
    | name       | メニュー種別 | 商品名 |
    | quantity   | 件数        | 個数   |
    | unit_price | 単価        | 単価   |
    | unit_cost  | 原価単価     | 原価   |

    **「原価単価」と「原価」は、どちらも1つあたりの値段。**
    メニュー側の見出しだけが「単価」まで書いてある。合計ではない。

    **行番号だけでは一意にならない。**2枚のシートから来るので、
    どちらの2行目かが決まらない。kind と組で一意にする。
    """

    MENU = "メニュー"
    PRODUCT = "商品"
    KIND_CHOICES = [(MENU, "メニュー"), (PRODUCT, "商品")]

    kind = models.CharField("種別", max_length=16, choices=KIND_CHOICES, db_index=True)
    sheet_row = models.IntegerField("シートの行", db_index=True)

    recorded_on = models.DateField("記録日", null=True, blank=True, db_index=True)
    name = models.CharField("名前", max_length=255)

    # 空欄は None。0（無料・サービス）と、記録していない、を混ぜない。
    quantity = models.IntegerField("数", null=True, blank=True)

    # **単価は整数ではない。**まとめ買いの集計行があるため。
    # 例: 天然だし調味粉の32〜35行は「3個で6995円」で、
    #     単価が 6995÷3 = 2331.6666… で入っている（2026-08-17 の書き出しで判明）。
    # 整数で受けると 2331 に切り捨てられ、3×2331=6993 で
    # **1行あたり2円、4行で8円が消える。**金額を勝手に減らさない。
    #
    # 小数6桁まで持てば、掛け戻して円に丸めたときに元の金額に戻る
    # （2331.666667 × 3 = 6995.000001 → 6995円）。
    unit_price = models.DecimalField(
        "単価（円）", max_digits=12, decimal_places=6, null=True, blank=True
    )
    # **メニューの112件は、原価単価がすべて 0。**空欄ではなく実際に0。
    # 「まだ入れていない0」ではなく「本当に原価0円」だと院長に確認済み
    # （2026-08-17）。教室や施術は仕入れが無いので、売上がそのまま残る。
    # 商品側は原価が入っている（例: よもぎ茶 1875円）。
    #
    # したがって 0 と空欄を区別する。**混ぜると粗利が変わる。**
    # 0 なら粗利＝売上、空欄なら粗利は出せない（None）。
    #
    # いまは全件が整数だが、単価と同じ理由でいつ割り算の結果が
    # 入ってもおかしくないので、こちらも小数で受ける。
    unit_cost = models.DecimalField(
        "原価（円・1つあたり）", max_digits=12, decimal_places=6, null=True, blank=True
    )

    memo = models.TextField("メモ", blank=True)

    deleted = models.BooleanField("削除済み", default=False, db_index=True)
    deleted_at = models.DateTimeField("削除日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "売上"
        verbose_name_plural = "売上"
        ordering = ["-recorded_on", "kind", "sheet_row"]
        constraints = [
            models.UniqueConstraint(
                fields=["kind", "sheet_row"], name="売上は種別と行で一意"
            )
        ]

    def __str__(self):
        return f"{self.recorded_on} {self.name}"

    @staticmethod
    def _円に丸める(数, 単価):
        """数×単価を、円（整数）にして返す。

        単価は小数を持つ（まとめ買いの集計行があるため）。
        掛けたあとに四捨五入すれば、元の金額に戻る。
        **先に単価を丸めてから掛けてはいけない。**そこで桁が落ちる。
        """
        if 数 is None or 単価 is None:
            return None
        return int((Decimal(数) * 単価).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

    @property
    def amount(self):
        """売れた金額（円）。数と単価がそろっている行だけ出せる。

        **片方が欠けていたら 0 ではなく None を返す。**
        0円売れたのと、分からないのは違う。
        """
        return self._円に丸める(self.quantity, self.unit_price)

    @property
    def cost(self):
        """かかった金額（円）。原価が空の行は None（0ではない）。"""
        return self._円に丸める(self.quantity, self.unit_cost)

    @property
    def profit(self):
        """残った金額。売上か原価のどちらかが分からなければ None。"""
        a, c = self.amount, self.cost
        if a is None or c is None:
            return None
        return a - c

    def to_dict(self):
        return {
            "kind": self.kind,
            "row": self.sheet_row,
            "date": self.recorded_on.isoformat() if self.recorded_on else None,
            "name": self.name,
            "quantity": self.quantity,
            "unitPrice": self.unit_price,
            "unitCost": self.unit_cost,
            "amount": self.amount,
            "memo": self.memo,
            "deleted": self.deleted,
        }
