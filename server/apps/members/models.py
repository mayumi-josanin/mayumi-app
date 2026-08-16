from django.db import models


class Member(models.Model):
    """会員台帳。スプレッドシートの「会員データ」にあたる。

    **すべての中心。**計測記録もアンケート回答も、ここに会員IDで結ぶ。

    設計で決めたこと（docs/design/移行設計.md）:

      ・会員IDは**いまのものをそのまま使う**。変えるとお客様が全員入り直しになる。
      ・「1人に1つ」のものだけをここに置く。何度も起きるもの（計測・回答）は別の表。
      ・パスコードは**平文で持たない。**いまスプレッドシートには平文で入っており、
        移すときがそれを直す機会になる。

    JSON のまま持っている列が3つある（特典履歴・スタンプ履歴・端末セッション）。
    設計では行に開くことにしているが、開くのは別の段階にする。
    いま無理に開くと、開き方を間違えたときに元が分からなくなる。
    まず**そのまま持ち込んで失わないこと**を優先し、開くのはあとから行う。
    """

    # スプレッドシートの「ID」。MYM-0000 の形。**変えない。**
    member_id = models.CharField("会員ID", max_length=32, primary_key=True)

    created_at = models.DateTimeField("登録日時", null=True, blank=True)

    name = models.CharField("氏名", max_length=100, db_index=True)
    kana = models.CharField("フリガナ", max_length=100, blank=True, db_index=True)
    phone = models.CharField("電話番号", max_length=32, blank=True)
    birthday = models.DateField("生年月日", null=True, blank=True)
    address = models.CharField("住所", max_length=255, blank=True)
    avatar_url = models.TextField("アイコンURL", blank=True)

    push_enabled = models.BooleanField("Push設定", default=False)
    status = models.CharField("ご状況", max_length=100, blank=True)

    # まゆみのスタンプカード。**ビジリスの回数券とは別物。**混ぜない。
    stamp_count = models.IntegerField("現在スタンプ数", default=0)
    stamp_card_number = models.IntegerField("スタンプカード番号", default=0)
    last_stamp_at = models.DateTimeField("最終スタンプ取得日時", null=True, blank=True)

    # ---- ログインに使うもの ----
    #
    # 平文のパスコードは持ち込まない。取り込むときにハッシュにする。
    # 4桁なので、漏れたときに総当たりで割られる。ハッシュだけでは足りないため、
    # ログインの窓口には呼び出しの回数制限（1分10回）をかけてある。
    passcode_hash = models.CharField("パスコード（ハッシュ）", max_length=255, blank=True)

    # 旧方式でご登録の2名ぶん。移行後にパスコードへ統合する。
    password_hash = models.CharField("パスワードハッシュ", max_length=255, blank=True)
    password_salt = models.CharField("パスワードソルト", max_length=255, blank=True)

    role = models.CharField("権限", max_length=32, blank=True)

    transfer_code = models.CharField("引き継ぎコード", max_length=64, blank=True)
    transfer_code_issued_at = models.DateTimeField("引き継ぎコード発行日時", null=True, blank=True)

    # ---- そのまま持ち込むもの（あとで行に開く）----
    device_sessions = models.JSONField("端末セッション", default=list, blank=True)
    stamp_history = models.JSONField("スタンプ履歴", default=list, blank=True)
    reward_history = models.JSONField("特典履歴", default=list, blank=True)

    # ---- 状態 ----
    #
    # 消さずに印を付ける。消してしまうと、その方の計測記録やアンケート回答が
    # 行き場を失う。
    #
    # 「削除状態」と「削除日時」は別々に持つ。
    # 印だけあって日時が無い方がいるため、日時に無理やり寄せると
    # 実在しない日付を作ることになる。**無い日時は無いままにする。**
    deleted = models.BooleanField("退会", default=False, db_index=True)
    deleted_at = models.DateTimeField("削除日時", null=True, blank=True)
    merged_into_id = models.CharField("統合先会員ID", max_length=32, blank=True)

    registration_source = models.CharField("登録経路", max_length=100, blank=True)
    registration_source_detail = models.CharField("登録経路詳細", max_length=255, blank=True)
    registration_source_updated_at = models.DateTimeField("登録経路更新日時", null=True, blank=True)

    last_online_at = models.DateTimeField("最終オンライン日時", null=True, blank=True)
    bijiris_registered = models.BooleanField("ビジリス登録", default=False)

    # この行がDBに入った・変わった時刻。移行の作業記録として持つ。
    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "会員"
        verbose_name_plural = "会員"
        ordering = ["member_id"]
        indexes = [models.Index(fields=["name", "kana"])]

    def __str__(self):
        return f"{self.member_id} {self.name}"

    @property
    def is_deleted(self):
        return self.deleted or self.deleted_at is not None

    def to_dict(self, 個人情報を含める=False):
        """GAS やアプリに返す形。

        **既定では、ログインに関わるものと連絡先を出さない。**
        必要な画面だけが `個人情報を含める=True` で取る。
        「とりあえず全部返す」を既定にすると、いつか要らない画面から漏れる。
        """
        def 日時(v):
            return v.isoformat() if v else None

        基本 = {
            "memberId": self.member_id,
            "name": self.name,
            "kana": self.kana,
            "stampCount": self.stamp_count,
            "stampCardNumber": self.stamp_card_number,
            "lastStampAt": 日時(self.last_stamp_at),
            "avatarUrl": self.avatar_url,
            "pushEnabled": self.push_enabled,
            "bijirisRegistered": self.bijiris_registered,
            "role": self.role,
            "deleted": self.is_deleted,
        }
        if not 個人情報を含める:
            return 基本

        基本.update({
            "phone": self.phone,
            "birthday": self.birthday.isoformat() if self.birthday else None,
            "address": self.address,
            "status": self.status,
            "createdAt": 日時(self.created_at),
            "lastOnlineAt": 日時(self.last_online_at),
            "registrationSource": self.registration_source,
            "registrationSourceDetail": self.registration_source_detail,
            "mergedIntoId": self.merged_into_id,
            "deviceSessions": self.device_sessions,
            "stampHistory": self.stamp_history,
            "rewardHistory": self.reward_history,
        })
        return 基本
