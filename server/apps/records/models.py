from decimal import ROUND_HALF_UP, Decimal

from django.db import models


class BackupRecord(models.Model):
    """控えの記録。スプレッドシートの BACKUP_LOG にあたる。139行 × 5列。

    GASがスプレッドシートの控え（コピー）を作ったときの記録。
    **控えそのものではなく、「いつ・何という名前で作ったか」の一覧。**
    実体はGoogleドライブにあり、ファイルIDとURLで辿れる。

    | 列 | 中身 |
    |---|---|
    | 作成日時 | 例: 2026-04-05T18:25:26+09:00 |
    | 種別 | 例: manual-admin（手で作ったか、自動か） |
    | ファイル名 | 例: まゆみ助産院_管理_backup_20260405_182523 |
    | ファイルID | ドライブ上のID |
    | URL | 開くためのリンク |

    **ドライブ上のファイルが消えても、この記録は残る。**
    「あの日の控えはもう無い」と分かること自体に意味がある。
    記録が消えると、無かったのか消えたのかも分からなくなる。

    設計書には130件とあったが、実物は139行（2026-08-17 の下見）。
    """

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    created_at = models.DateTimeField("作成日時", null=True, blank=True, db_index=True)
    kind = models.CharField("種別", max_length=64, blank=True, db_index=True)

    file_name = models.CharField("ファイル名", max_length=255, blank=True)
    # ドライブのIDは33文字前後だが、余裕を持たせる。
    file_id = models.CharField("ファイルID", max_length=128, blank=True, db_index=True)
    url = models.TextField("URL", blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)

    class Meta:
        verbose_name = "控えの記録"
        verbose_name_plural = "控えの記録"
        ordering = ["-created_at", "-sheet_row"]

    def __str__(self):
        return f"{self.created_at} {self.file_name}"


class AuditLog(models.Model):
    """操作履歴。スプレッドシートの ADMIN_AUDIT_LOG にあたる。

    **9,373行のうち、移すのは「人の操作」794件だけ**（院長の判断）。
    残り8,579件（91.5%）はアプリが勝手に行う自動処理の記録で、
    誰も見ない。全部移すとデータ量が11倍になり、控えも大きくなる。

    除いた2種（2026-08-17 に全行を数えた結果）:

    | 種別 | 件数 | 割合 | 中身 |
    |---|---|---|---|
    | syncUserDeviceSession | 7,788 | 83.1% | 端末の記録合わせ |
    | syncUserRewardStatus | 791 | 8.4% | 特典状態の同期 |

    **除いた件数は取り込みのたびに数えて出す。**黙って減らさない。

    「操作者」は長らく全件「管理者」で、誰がやったかは記録されていなかった。
    **ただし2026-08-17から「ご本人」が入り始めている**（お客様自身の
    パスコード再設定）。移した794件のうち2件がそれ。
    決め打ちにせず、入ってきた値をそのまま持つ。

    **「概要」は種別によって埋まり方がまったく違う。**
    先頭200行だけ見ると 0/199 で空に見えるが、全体では
    deleteOrders 30/30、mergeUsers 13/13 と全部埋まっている種別がある。
    空欄のまま持ち、**「使っていない列」と決めつけない。**
    """

    # 自動処理の記録。これは移さない。**増やすときは移行設計にも書くこと。**
    自動処理の種別 = ("syncUserDeviceSession", "syncUserRewardStatus")

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    happened_at = models.DateTimeField("日時", null=True, blank=True, db_index=True)
    kind = models.CharField("種別", max_length=64, db_index=True)

    # 成功 / 失敗
    result = models.CharField("結果", max_length=32, blank=True, db_index=True)

    target = models.TextField("対象", blank=True)
    summary = models.TextField("概要", blank=True)

    # 「管理者」または「ご本人」。長らく「管理者」だけだった。
    operator = models.CharField("操作者", max_length=64, blank=True, db_index=True)

    # 何を送って何が返ったかの控え。読める形（辞書）で持つ。
    # 文字列のままだと、あとで中を絞り込めない。
    detail = models.JSONField("詳細", null=True, blank=True)
    # 読めなかったときは、元の文字列を捨てずにここへ置く。
    # **形が違うからといって記録を消さない。**
    detail_raw = models.TextField("詳細（読めなかったもの）", blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)

    class Meta:
        verbose_name = "操作履歴"
        verbose_name_plural = "操作履歴"
        ordering = ["-happened_at", "-sheet_row"]
        indexes = [models.Index(fields=["kind", "-happened_at"])]

    def __str__(self):
        return f"{self.happened_at} {self.kind} {self.result}"

    @property
    def failed(self):
        return self.result == "失敗"


class SupplierPrice(models.Model):
    """仕入値。スプレッドシートの「管理マスタ」にあたる。22行 × 5列。

    **5列のうち、使うのは3列だけ**（2026-08-17 の下見）:

    | 列 | 中身 |
    |---|---|
    | 1. 商品名（完全一致） | 22件 |
    | 2. 仕入値（円） | 22件 |
    | 3. 備考 | 14件 |
    | 4. （見出しなし） | **0件。空の列** |
    | 5. 【仕入値の自動入力について】 | **3件。人向けの説明文** |

    5列目は「注文管理シートでは、D列の『商品名』とここの『商品名』を
    照合して…」という運用メモが3行だけ入っている。**列として移すと、
    仕入値の表に意味のない文字列が混ざる。**見出しの名前で拾い、
    この2列は取らない。

    **商品名で商品マスタと結びついている。**見出しが「商品名（完全一致）」で、
    名前がぴたり合うことを前提にした作り。IDに直したくなるが、直すと
    名前が少し違うだけの商品が迷子になる。移行のあいだは名前のまま置き、
    **結びついたかどうかを記録して、結びつかないものは数えて報告する。**
    """

    sheet_row = models.IntegerField("シートの行", db_index=True, unique=True)

    product_name = models.CharField("商品名（完全一致）", max_length=255, db_index=True)

    # 単価と同じ理由で小数で持つ。いまは全件が整数だが、
    # 割り算の結果（3個で◯◯円 など）が入ってもおかしくない。
    price = models.DecimalField(
        "仕入値（円）", max_digits=12, decimal_places=6, null=True, blank=True
    )

    memo = models.TextField("備考", blank=True)

    # 商品マスタに同じ名前があったか。取り込んだときの結果を残す。
    # **見つからないことを黙って通さない。**あとで粗利を出すときに、
    # どの商品の原価が分からないのかが即座に分かる。
    matched_product = models.ForeignKey(
        "content.Product", verbose_name="結びついた商品",
        null=True, blank=True, on_delete=models.SET_NULL,
        related_name="supplier_prices",
    )

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "仕入値"
        verbose_name_plural = "仕入値"
        ordering = ["sheet_row"]

    def __str__(self):
        return f"{self.product_name} {self.price}円"

    @property
    def is_linked(self):
        return self.matched_product_id is not None


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


# 回数券分析は、お客様のお体についての記録なので別ファイルに置いている。
# ここで読み込むことで Django が表として認識する。
from .models_ticket import TicketAnalysis  # noqa: E402,F401


class AppSetting(models.Model):
    """アプリの設定。GAS のスクリプトプロパティにあたる。

    表ではなくプロパティに入っていたもの。`getAppRuntimeConfig` と
    `getRewardGachaConfig` がここから読む。

    **鍵と値だけの素直な作りにする。**中身の形は設定ごとに違う
    （片方はバージョン情報、もう片方は4か月分の景品表）ので、
    列に開かずJSONのまま持つ。GASも `JSON.stringify` で入れている。

    移したもの（2026-08-23）:

    | 鍵 | 中身 |
    |---|---|
    | `APP_RUNTIME_CONFIG` | アプリの版・更新案内の文言 |
    | `REWARD_GACHA_CONFIG` | 4か月分の景品と確率 |

    **秘密は入れない。**`ADMIN_TOKEN_SECRET` などは `.env` にあり、
    ここには置かない。置くと、控え（バックアップ）に秘密が混ざる。
    """

    key = models.CharField("鍵", max_length=64, primary_key=True)
    value = models.JSONField("値", default=dict, blank=True)

    note = models.CharField("覚え書き", max_length=255, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "アプリの設定"
        verbose_name_plural = "アプリの設定"
        ordering = ["key"]

    def __str__(self):
        return self.key
