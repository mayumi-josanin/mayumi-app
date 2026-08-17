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


class SupportFaq(models.Model):
    """APP_SUPPORT_FAQ。使い方チャットが答えるための質問と回答。
    46行 × 7列（2026-08-17 時点）。全列が46件で、欠けている項目は無い。

    **掲載の共通を継承しない。**削除状態・公開開始日時などの列がそもそも無い。
    無い列を「あることにして」表を作ると、空の欄が意味ありげに残る。

    キーワードは「プロフィール,登録,会員,…」のような読点区切りの文字列。
    GAS側が `[\\n,、，/空白]` で分けて点数を付けている。**分け方は
    答えの当たり外れそのもの**なので、分けずに文字列のまま移し、
    分ける処理はサーバー側でも同じ規則で書く。
    """

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    published = models.BooleanField("公開", default=False, db_index=True)
    category = models.CharField("カテゴリ", max_length=100, blank=True, db_index=True)
    question = models.CharField("質問", max_length=255)
    keywords = models.TextField("キーワード", blank=True)
    answer = models.TextField("回答", blank=True)

    # 大きいほど先に出る（例: 154）。同点はシートの行の順。
    priority = models.IntegerField("優先度", default=0)

    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "使い方FAQ"
        verbose_name_plural = "使い方FAQ"
        ordering = ["-priority", "sheet_row"]

    def __str__(self):
        return self.question

    @property
    def keyword_list(self):
        """GAS の splitSupportKeywords_ と同じ分け方。**両方そろえること。**"""
        import re

        return [x for x in re.split(r"[\n,、，/\s]+", self.keywords or "") if x.strip()]

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "status": "公開" if self.published else "非公開",
            "category": self.category,
            "question": self.question,
            "keywords": self.keywords,
            "answer": self.answer,
            "priority": self.priority,
        }


class PushNotice(models.Model):
    """PUSH_NOTICES。送ったお知らせ通知の控え。96行 × 15列（2026-08-17 時点）。

    シートの見出しは13列だが、実物は15列ある。14・15列目の
    削除状態・削除日時は、あとから汎用の削除の仕組みが足したもので、
    どちらも0件。**実物を見なければ2列を落としていた。**

    **日時・通知ID・送信結果の3つは、揃って埋まるか揃って空。**
    96件中42件だけが埋まっている。ステータス別に数えると
    「自動送信済み」90件中40件、「送信済み」6件中2件で、
    ステータスとは関係が無かった。OneSignal から通知IDが返ってきた
    配信がこの42件で、残りは結果を受け取らずに記録だけ残った行。
    **空欄は空欄のまま移す。**0や仮の日時で埋めると、
    「いつ届いたか」が分からなくなるどころか、届いたことになってしまう。
    """

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    # 実際に送られた日時。届かなかった行は空のまま。
    sent_at = models.DateTimeField("日時", null=True, blank=True, db_index=True)
    title = models.CharField("タイトル", max_length=255)
    body = models.TextField("本文", blank=True)

    # 誰に送ったか。実物は96件すべて 'all'。
    target_status = models.CharField("送信対象", max_length=100, blank=True)
    # GAS の作りでは、会員を選んで送ると**会員の一覧（JSON）が入りうる。**
    # 書き出す前に6件の中身を確かめたところ、実際に入っていたのは
    # 「テスト送信:デモ2」「確認送信:デモ2」というラベルだけで、
    # 会員は含まれていなかった（2026-08-17）。
    # **将来ここに会員一覧が入ったら、書き出したJSONは個人情報を含む。**
    # そのときは 書き出したJSONを片付ける() の要注意扱いに加えること。
    target_detail = models.TextField("送信対象詳細", blank=True)
    recipient_count = models.IntegerField("送信件数", null=True, blank=True)

    # 自動送信済み / 送信済み / 下書き / 予約済み / 送信失敗
    status = models.CharField("ステータス", max_length=32, blank=True, db_index=True)

    # 予約配信。実物は96件すべて空で、いまのところ使われていない。
    scheduled_at = models.DateTimeField("配信予定日時", null=True, blank=True)

    # 通知を押したときに開くページ（home / shop / calendar / news / …）。
    target_page = models.CharField("対象ページ", max_length=32, blank=True)
    preview_body = models.TextField("プレビュー本文", blank=True)

    notification_id = models.CharField("通知ID", max_length=64, blank=True)
    result = models.TextField("送信結果", blank=True)

    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    deleted = models.BooleanField("削除済み", default=False, db_index=True)
    deleted_at = models.DateTimeField("削除日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "プッシュ通知"
        verbose_name_plural = "プッシュ通知"
        ordering = ["-sheet_row"]

    def __str__(self):
        return f"{self.sent_at or '（未送信）'} {self.title}"

    @property
    def was_delivered(self):
        """本当に配信されたと言い切れるか。通知IDが返ってきた行だけ。"""
        return bool(self.notification_id)

    def to_dict(self):
        return {
            "row": self.sheet_row,
            "sentAt": self.sent_at.isoformat() if self.sent_at else None,
            "title": self.title,
            "body": self.body,
            "targetStatus": self.target_status,
            "recipientCount": self.recipient_count,
            "status": self.status,
            "targetPage": self.target_page,
            "notificationId": self.notification_id,
            "deleted": self.deleted,
        }


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
