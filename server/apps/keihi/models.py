"""経費管理のモデル。現金出納帳（JDL 会計伝票 様式 777）とレシート。

**差引残高と合計は持たない。** 残高は行の並びで決まるので、持つと必ずズレる。
表示のたびに services.py で計算する。
"""

from django.db import models


class 時刻付き(models.Model):
    """作成・更新の日時。apps/dev と同じ作り。"""

    created_at = models.DateTimeField("作成日時", auto_now_add=True)
    updated_at = models.DateTimeField("更新日時", auto_now=True)

    class Meta:
        abstract = True


class Cashbook(時刻付き):
    """帳簿。年月ごとに1冊。紙の様式の1ページにあたる。"""

    year = models.IntegerField("年")
    month = models.IntegerField("月")
    opening_balance = models.IntegerField("前葉繰越", default=0)

    class Meta:
        verbose_name = "現金出納帳"
        verbose_name_plural = "現金出納帳"
        constraints = [
            models.UniqueConstraint(fields=["year", "month"], name="uq_cashbook_year_month"),
        ]
        ordering = ["-year", "-month"]

    def __str__(self):
        return f"{self.year}年{self.month}月"


class CashbookEntry(時刻付き):
    """出納帳の1行。入力するのは5項目だけ。"""

    book = models.ForeignKey(Cashbook, on_delete=models.CASCADE, related_name="entries", verbose_name="帳簿")
    row_order = models.IntegerField("並び順")
    month = models.IntegerField("日付（月）", null=True, blank=True)
    day = models.IntegerField("日付（日）", null=True, blank=True)
    description = models.TextField("摘要", blank=True, default="")
    counter_account = models.CharField("相手科目", max_length=100, blank=True, default="")
    income = models.IntegerField("収入金額", null=True, blank=True)
    payment = models.IntegerField("支払金額", null=True, blank=True)

    class Meta:
        verbose_name = "現金出納帳の行"
        verbose_name_plural = "現金出納帳の行"
        ordering = ["row_order", "id"]

    def __str__(self):
        return f"{self.book} {self.row_order}行目"


class Receipt(時刻付き):
    """レシート1枚。写真を読み取った結果を置き、確かめてから出納帳の行にする。

    **読み取った時点では出納帳に入れない。** AI は金額の桁や日付を読み違えることがあり、
    そのまま帳簿に入ると、あとから探すのが大変になる。院長が見て「出納帳へ」を押したものだけが行になる。
    """

    状態 = [
        ("pending", "読み取り待ち"),
        ("done", "読み取り済み"),
        ("failed", "読み取れなかった"),
        ("manual", "手入力"),
    ]
    種別 = [("payment", "支払"), ("income", "収入")]

    # 写真。手入力の1件は空。**公開の /media/ では配らない**（金額と店名が写っているため）。
    # 管理画面にログインした人だけが見られる専用の配り口から出す（views.receipt_image）
    image_name = models.CharField("写真のファイル名", max_length=200, blank=True, default="")
    original_filename = models.CharField("元のファイル名", max_length=500, blank=True, default="")

    date = models.DateField("日付", null=True, blank=True)
    amount = models.IntegerField("金額", null=True, blank=True)
    store_name = models.CharField("店名・支払先", max_length=255, blank=True, default="")
    counter_account = models.CharField("相手科目", max_length=100, blank=True, default="")
    kind = models.CharField("種別", max_length=10, choices=種別, default="payment")
    memo = models.TextField("メモ", blank=True, default="")

    status = models.CharField("状態", max_length=10, choices=状態, default="pending")
    error = models.TextField("読み取れなかった理由", blank=True, default="")
    raw = models.TextField("AI の返事（そのまま）", blank=True, default="")

    # 出納帳へ反映したときの行。行が消えたらここは空になる（レシート自体は残す）
    entry = models.ForeignKey(
        CashbookEntry, on_delete=models.SET_NULL, null=True, blank=True,
        related_name="receipts", verbose_name="出納帳の行",
    )

    class Meta:
        verbose_name = "レシート"
        verbose_name_plural = "レシート"
        # 既定は日付の古い順。帳簿は起きた順に並べて見るものなので
        ordering = ["date", "id"]

    def __str__(self):
        return f"{self.date or '日付なし'} {self.store_name or '（店名なし）'}"

    @property
    def 出納帳へ反映済み(self) -> bool:
        return self.entry_id is not None
