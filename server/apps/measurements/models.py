from django.db import models


class Measurement(models.Model):
    """ビジリスの計測記録。スプレッドシートの「測定履歴」にあたる。

    1人が何回も測るので、会員台帳とは別の表にする。

    member_id は空を許す。アンケートの送信には会員登録が要らないため、
    「記録はあるが会員ではない方」がいつでも発生しうる。
    必須にすると、その方の記録を移せなくなる。
    customer_name は常に残し、あとから会員が見つかれば member_id を埋める。
    """

    # スプレッドシートの「測定ID」。同じ人・同じ日は同じIDになるので、
    # 何度取り込んでも二重に増えない。
    measurement_id = models.CharField("測定ID", max_length=64, unique=True)

    member_id = models.CharField("会員ID", max_length=32, blank=True, db_index=True)
    customer_name = models.CharField("お名前", max_length=100, db_index=True)
    member_number = models.CharField("会員番号", max_length=32, blank=True)

    measured_on = models.DateField("測定日", db_index=True)

    waist = models.DecimalField("ウエスト(cm)", max_digits=5, decimal_places=1, null=True, blank=True)
    hip = models.DecimalField("ヒップ(cm)", max_digits=5, decimal_places=1, null=True, blank=True)
    thigh_right = models.DecimalField("太もも右(cm)", max_digits=5, decimal_places=1, null=True, blank=True)
    thigh_left = models.DecimalField("太もも左(cm)", max_digits=5, decimal_places=1, null=True, blank=True)
    whr = models.DecimalField("WHR", max_digits=5, decimal_places=3, null=True, blank=True)

    staff_memo = models.TextField("スタッフメモ", blank=True)

    # スプレッドシート側の日時をそのまま持ち込む（移行の突き合わせに使う）。
    created_at = models.DateTimeField("作成日時", null=True, blank=True)
    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    # この行がDBに入った・変わった時刻。移行の作業記録として持つ。
    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "計測記録"
        verbose_name_plural = "計測記録"
        ordering = ["customer_name", "measured_on"]
        indexes = [models.Index(fields=["customer_name", "measured_on"])]

    def __str__(self):
        return f"{self.customer_name} {self.measured_on}"

    def to_dict(self):
        """GAS がいま扱っている形に合わせて返す。

        アプリ側の見え方を変えないため、キーはスプレッドシートの見出しに寄せる。
        """
        def num(value):
            return float(value) if value is not None else None

        return {
            "measurementId": self.measurement_id,
            "memberId": self.member_id,
            "customerName": self.customer_name,
            "memberNumber": self.member_number,
            "measuredOn": self.measured_on.isoformat(),
            "waist": num(self.waist),
            "hip": num(self.hip),
            "thighRight": num(self.thigh_right),
            "thighLeft": num(self.thigh_left),
            "whr": num(self.whr),
            "staffMemo": self.staff_memo,
            "createdAt": self.created_at.isoformat() if self.created_at else None,
            "updatedAt": self.updated_at.isoformat() if self.updated_at else None,
        }
