from django.db import models


class 掲載の共通(models.Model):
    """お知らせもカレンダーも共通で持つ「掲載の状態」。

    どちらのシートも、後から列が足されて同じ考え方に育っている。
    ・公開設定       … 公開 / 非公開
    ・お知らせ一覧公開 … アプリの「お知らせ」に出すか
    ・削除状態・削除日時 … 消さずに印を付ける（論理削除）
    ・公開開始日時   … 予約投稿

    **消さずに印を付ける。**消してしまうと、いつ何を出していたかが
    分からなくなる。お客様からの問い合わせは過去の掲載についても来る。
    """

    # シートの行番号。移行のあいだ、元の行と突き合わせるために持つ。
    # 会員IDのような安定した鍵が無い表なので、これが手がかりになる。
    sheet_row = models.IntegerField("シートの行", db_index=True)

    published = models.BooleanField("公開", default=False, db_index=True)
    notice_listed = models.BooleanField("お知らせ一覧に出す", default=False)

    deleted = models.BooleanField("削除済み", default=False, db_index=True)
    deleted_at = models.DateTimeField("削除日時", null=True, blank=True)
    delete_reason = models.CharField("削除理由", max_length=255, blank=True)

    publish_at = models.DateTimeField("公開開始日時", null=True, blank=True)
    notice_listed_at = models.DateTimeField("お知らせ一覧掲載日時", null=True, blank=True)
    notice_delisted_at = models.DateTimeField("お知らせ一覧削除日時", null=True, blank=True)

    sort_order = models.IntegerField("表示順", null=True, blank=True)
    image_url = models.TextField("画像URL", blank=True)

    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        abstract = True

    @property
    def is_visible(self):
        """いまお客様に見えている状態か。"""
        return self.published and not self.deleted


class News(掲載の共通):
    """ブログ・お知らせ。スプレッドシートの「ブログ・お知らせ」にあたる。

    95行 × 20列（2026-08-17 時点）。設計書には94件とあったが、実物は95行。
    **列はコードから推測せず、実物の1行目を見て決めた。**
    """

    posted_on = models.DateField("投稿日", null=True, blank=True, db_index=True)
    title = models.CharField("タイトル", max_length=255)
    category = models.CharField("カテゴリ", max_length=100, blank=True, db_index=True)
    icon = models.CharField("アイコン絵文字", max_length=16, blank=True)
    body = models.TextField("本文", blank=True)

    # 「Instagramはこちら」のような、記事の下に置くボタン。
    link_url = models.TextField("リンクURL", blank=True)
    link_label = models.CharField("リンクボタン名", max_length=100, blank=True)
    button_text = models.CharField("ボタンテキスト", max_length=100, blank=True)

    class Meta:
        verbose_name = "お知らせ"
        verbose_name_plural = "お知らせ"
        ordering = ["-posted_on", "-sheet_row"]

    def __str__(self):
        return f"{self.posted_on} {self.title}"

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "date": self.posted_on.isoformat() if self.posted_on else None,
            "title": self.title,
            "category": self.category,
            "icon": self.icon,
            "body": self.body,
            "imageUrl": self.image_url,
            "linkUrl": self.link_url,
            "linkLabel": self.link_label,
            "buttonText": self.button_text,
            "published": self.published,
            "noticeListed": self.notice_listed,
            "deleted": self.deleted,
        }


class Menu(掲載の共通):
    """MENUS。院のメニュー（教室・施術）。13行 × 16列（2026-08-17 時点）。

    **画像URLが配列で入っている。**1つのメニューに複数の画像が付く。
    文字列のまま持つと、あとで1枚ずつ扱えない。読める形（配列）で持つ。

    **表示順が13桁の数値。**日時をそのまま並び順に使っている
    （例: 1775290672114 = ミリ秒）。ふつうの整数では収まらないので
    大きい型で持つ。並び順の意味は変えずにそのまま移す。
    """

    registered_on = models.DateField("登録日", null=True, blank=True)
    name = models.CharField("メニュー名", max_length=255)
    summary = models.TextField("概要説明", blank=True)
    category = models.CharField("カテゴリ", max_length=100, blank=True, db_index=True)
    booking_status = models.CharField("予約状況", max_length=100, blank=True)

    # 複数枚。空なら []。
    image_urls = models.JSONField("画像URL", default=list, blank=True)

    # 掲載の共通が持つ sort_order は整数だが、こちらは13桁が入るため別に持つ。
    sort_key = models.BigIntegerField("表示順", null=True, blank=True)

    class Meta:
        verbose_name = "メニュー"
        verbose_name_plural = "メニュー"
        ordering = ["-sort_key", "sheet_row"]

    def __str__(self):
        return self.name

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "name": self.name,
            "summary": self.summary,
            "category": self.category,
            "bookingStatus": self.booking_status,
            "imageUrls": self.image_urls,
            "registeredOn": self.registered_on.isoformat() if self.registered_on else None,
            "published": self.published,
            "noticeListed": self.notice_listed,
            "deleted": self.deleted,
        }


class Product(掲載の共通):
    """商品マスタ。9行 × 16列（2026-08-17 時点）。

    在庫を持つ。**在庫は数えられる状態のまま移す。**
    「0個」と「そもそも在庫を管理していない」は別の意味なので、
    空欄は None にして、0 と区別する。
    """

    name = models.CharField("商品名", max_length=255)
    category = models.CharField("カテゴリ", max_length=100, blank=True, db_index=True)
    price = models.IntegerField("価格（円）", null=True, blank=True)
    special_price = models.IntegerField("特別価格（円）", null=True, blank=True)

    description = models.TextField("商品説明", blank=True)
    icon_url = models.TextField("アイコン", blank=True)
    description_image_url = models.TextField("商品説明画像", blank=True)
    background_color = models.CharField("背景色コード", max_length=32, blank=True)

    # 空欄は「管理していない」。0（売り切れ）と区別する。
    stock = models.IntegerField("在庫数", null=True, blank=True)
    stock_warning = models.IntegerField("在庫警告閾値", null=True, blank=True)

    class Meta:
        verbose_name = "商品"
        verbose_name_plural = "商品"
        ordering = ["sheet_row"]

    def __str__(self):
        return self.name

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "name": self.name,
            "category": self.category,
            "price": self.price,
            "specialPrice": self.special_price,
            "description": self.description,
            "iconUrl": self.icon_url,
            "descriptionImageUrl": self.description_image_url,
            "backgroundColor": self.background_color,
            "stock": self.stock,
            "stockWarning": self.stock_warning,
            "published": self.published,
            "deleted": self.deleted,
        }


class Category(models.Model):
    """カテゴリマスタ。7行 × 2列（2026-08-17 時点）。

    お知らせ・ブログ・メニュー・商品が、このカテゴリ名を文字列で参照している。
    **いまは名前で結びついている。**IDに直したくなるが、直すと参照元を
    すべて書き換えることになるので、移行のあいだは名前のまま置く。

    2列目に**見出しが入っていない。**値は「お知らせ」「ブログ」で、
    カテゴリの種別。名前で拾えないので、位置（2列目）で拾う。
    """

    sheet_row = models.IntegerField("シートの行", db_index=True)
    name = models.CharField("カテゴリ名", max_length=100, db_index=True)
    kind = models.CharField("種別", max_length=32, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "カテゴリ"
        verbose_name_plural = "カテゴリ"
        ordering = ["sheet_row"]

    def __str__(self):
        return f"{self.name}（{self.kind}）" if self.kind else self.name

    def to_dict(self):
        return {"row": self.sheet_row, "name": self.name, "kind": self.kind}


class CalendarEvent(掲載の共通):
    """カレンダー。スプレッドシートの「カレンダー」にあたる。

    143行 × 19列（2026-08-17 時点）。うち21件は削除済みの印が付いている。
    設計書には116件とあったが、それは削除済みを除いた数と思われる。
    **消さずに印ごと移す。**あとから「あの日は何をしていたか」を辿れるように。
    """

    event_on = models.DateField("日付", null=True, blank=True, db_index=True)
    title = models.CharField("イベント名", max_length=255)
    detail = models.TextField("詳細", blank=True)
    color = models.CharField("カラー", max_length=32, blank=True)

    # 「対象メニュー行」。メニュー表の行を指している。
    # メニューを移すときに、正しい結びつきへ直す。
    menu_row = models.IntegerField("対象メニュー行", null=True, blank=True)

    class Meta:
        verbose_name = "カレンダー"
        verbose_name_plural = "カレンダー"
        ordering = ["-event_on", "-sheet_row"]

    def __str__(self):
        return f"{self.event_on} {self.title}"

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "date": self.event_on.isoformat() if self.event_on else None,
            "title": self.title,
            "detail": self.detail,
            "color": self.color,
            "imageUrl": self.image_url,
            "menuRow": self.menu_row,
            "published": self.published,
            "noticeListed": self.notice_listed,
            "deleted": self.deleted,
        }
