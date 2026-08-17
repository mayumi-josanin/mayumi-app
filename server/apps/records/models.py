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
    unit_price = models.IntegerField("単価（円）", null=True, blank=True)
    # **メニューの112件は、原価単価がすべて 0。**空欄ではなく実際に0。
    # 「まだ入れていない0」ではなく「本当に原価0円」だと院長に確認済み
    # （2026-08-17）。教室や施術は仕入れが無いので、売上がそのまま残る。
    # 商品側は原価が入っている（例: よもぎ茶 1875円）。
    #
    # したがって 0 と空欄を区別する。**混ぜると粗利が変わる。**
    # 0 なら粗利＝売上、空欄なら粗利は出せない（None）。
    unit_cost = models.IntegerField("原価（円・1つあたり）", null=True, blank=True)

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

    @property
    def amount(self):
        """売れた金額。数と単価がそろっている行だけ出せる。

        **片方が欠けていたら 0 ではなく None を返す。**
        0円売れたのと、分からないのは違う。
        """
        if self.quantity is None or self.unit_price is None:
            return None
        return self.quantity * self.unit_price

    @property
    def cost(self):
        """かかった金額。原価が空の行は None（0ではない）。"""
        if self.quantity is None or self.unit_cost is None:
            return None
        return self.quantity * self.unit_cost

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
