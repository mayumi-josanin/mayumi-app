"""経費管理のモデル。いまは現金出納帳（JDL 会計伝票 様式 777）だけ。

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
