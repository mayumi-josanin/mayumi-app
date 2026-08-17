"""回数券分析の表だけを分けて置く。

records/models.py が長くなってきたので、お客様に紐づく記録はこちらへ。
（Django からは records/models.py で読み込む）
"""

from django.db import models


class TicketAnalysis(models.Model):
    """回数券分析結果。ビジリスのスプレッドシートの「回数券分析結果」にあたる。

    回数券を使い終えたお客様の、ビフォー・アフター写真をAIが見て
    書いた分析。**お客様のお体についての記録**なので、扱いは会員データと同じ。

    | 列 | 中身 |
    |---|---|
    | 作成日時 / 更新日時 | |
    | 回答ID | アンケート回答への手がかり |
    | お名前 | **会員番号は無い** |
    | 提出日時 | |
    | ビフォー写真JSON / アフター写真JSON | ドライブ上の写真の場所 |
    | 分析状態 | |
    | 分析結果 | AIが書いた本文 |
    | 分析日時 / エラー | |

    **会員には結びつけない。**この表には会員番号が無く、あるのはお名前だけ。
    CLAUDE.md にあるとおり、お名前で探すと同姓同名・改名・表記ゆれで
    別の方に行き着く（実際に記録が混ざったことがある）。
    お名前をそのまま持ち、`member_id` は空のままにしてある。

    あとから会員番号で結びつける道が要るときは、
    **お名前ではなく回答IDから辿る**こと。回答IDはアンケート回答に付いていて、
    そちらには会員番号が入っている。
    """

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    # 空のまま。**お名前から埋めない。**回答ID経由で確かめられたときだけ入れる。
    member_id = models.CharField("会員ID", max_length=32, blank=True, db_index=True)
    customer_name = models.CharField("お名前", max_length=255, blank=True, db_index=True)

    # アンケート回答への手がかり。会員に結びつけるならここから辿る。
    response_id = models.CharField("回答ID", max_length=128, blank=True, db_index=True)

    submitted_at = models.DateTimeField("提出日時", null=True, blank=True)

    # ドライブ上の写真の場所。読める形（配列）で持つ。
    # 読めなかったときは元の文字列を捨てずに残す。
    before_photos = models.JSONField("ビフォー写真", null=True, blank=True)
    after_photos = models.JSONField("アフター写真", null=True, blank=True)
    photos_raw = models.TextField("写真（読めなかったもの）", blank=True)

    status = models.CharField("分析状態", max_length=64, blank=True, db_index=True)
    result = models.TextField("分析結果", blank=True)
    analyzed_at = models.DateTimeField("分析日時", null=True, blank=True)
    error = models.TextField("エラー", blank=True)

    created_at = models.DateTimeField("作成日時", null=True, blank=True)
    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)

    class Meta:
        verbose_name = "回数券分析"
        verbose_name_plural = "回数券分析"
        ordering = ["-submitted_at", "-sheet_row"]

    def __str__(self):
        return f"{self.customer_name} {self.submitted_at}"

    @property
    def is_done(self):
        return bool(self.result) and not self.error
