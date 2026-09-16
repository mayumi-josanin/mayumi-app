"""ビジリスの表。ビジリスの GAS（bijiris/gas/Code.gs）が持っている形をそのまま写す。

| ビジリスの置き場 | ここ |
|---|---|
| プロパティ SURVEYS_JSON（`loadSurveys_` / `validateSurveyPayload_`） | Survey |
| シート「回答一覧」（MASTER_HEADERS）＋ アンケートごとのシート（`ensureSurveySheet_`） | Response |
| 回答JSON / 写真JSON の中の写真（`savePhotoFiles_` が Drive に置いたもの） | ResponsePhoto |
| プロパティ CUSTOMER_PROFILES_JSON（`normalizeCustomerProfileRecord_`）＋ CUSTOMER_MEMOS_JSON（`normalizeCustomerMemoRecord_`） | CustomerProfile |
| シート「回数券分析結果」 | records.TicketAnalysis（既存） |
| プロパティ ADMIN_PREFERENCES_JSON / TICKET_SURVEY_META_JSON / TICKET_SURVEY_PROMPT | records.AppSetting（鍵 bijiris_preferences / bijiris_ticket_meta / bijiris_ticket_prompt） |

**お名前で会員に結びつけない**（CLAUDE.md）。ビジリスの記録はどれも「お名前」が鍵で会員番号を持たない。
`member_id` は空のまま持ち、受付が画面で結んだときだけ入る。

**Code.gs の「読む関数」は書く。**`loadSurveys_` `getPreferences_` `getCustomerProfiles_` `getCustomerMemos_` は、
正規化した結果が元と違えばプロパティへ書き戻す（`getCustomerProfiles_` は会員番号の採番までする）。
写しを取るときはこれらを呼ばず、`PropertiesService.getScriptProperties().getProperty(鍵)` で生のまま読むこと
（gas/ビジリスを書き出す.js）。
"""

from django.db import models


class Survey(models.Model):
    """アンケートの定義。SURVEYS_JSON の1本ぶん（本番は2本: 施術後アンケート・計測時アンケート）。

    項目は `validateSurveyPayload_`（Code.gs 866行〜）が返す形そのまま。
    設問（questions）は JSON のまま持つ。1問の形:
        {id, label, type(text/textarea/rating/choice/checkbox/photo), required, options[],
         placeholder, visibilityConditions[{questionId, value}], visibleWhen{questionId, value}|null}
    `visibleWhen` は `visibilityConditions[0]` の写し（古いアプリ向け）。列に開くと2か所そろえる決まり
    （CLAUDE.md 6: default-surveys.js と Code.gs）が3か所になるので、開かない。
    """

    STATUS_CHOICES = [("published", "公開中"), ("draft", "下書き"), ("archived", "終了")]

    survey_id = models.CharField("アンケートID", max_length=128, unique=True)
    title = models.CharField("タイトル", max_length=255)
    description = models.TextField("説明文", blank=True)
    intro_message = models.TextField("回答前の案内", blank=True)
    completion_message = models.TextField("回答後の文言", blank=True)
    # normalizeSurveyStatus_: published / draft / archived。それ以外は published 扱い
    status = models.CharField("公開状態", max_length=16, choices=STATUS_CHOICES, default="published", db_index=True)
    sort_order = models.IntegerField("並び", default=0)
    accepting_responses = models.BooleanField("回答を受け付ける", default=True)
    start_at = models.DateTimeField("受付開始", null=True, blank=True)
    end_at = models.DateTimeField("受付終了", null=True, blank=True)
    questions = models.JSONField("設問", default=list, blank=True)

    created_at = models.DateTimeField("作成日時", null=True, blank=True)
    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "アンケート"
        verbose_name_plural = "アンケート"
        ordering = ["sort_order", "survey_id"]

    def __str__(self):
        return self.title

    @property
    def photo_question_ids(self):
        return [q.get("id") for q in (self.questions or []) if isinstance(q, dict) and q.get("type") == "photo"]

    @property
    def question_count(self):
        return len(self.questions or [])


class Response(models.Model):
    """回答1件。シート「回答一覧」の1行（MASTER_HEADERS）＋ アンケートごとのシートの1行。

    | 回答一覧の列 | ここ |
    |---|---|
    | 送信日時 | submitted_at |
    | 回答ID | response_id |
    | アンケートID / アンケート名 | survey_key / survey_title（survey は取り込み時に survey_id で結ぶ） |
    | 端末ID | client_id |
    | お名前 / メールアドレス | customer_name / customer_email |
    | 対応状況 | status（normalizeStatus_: new / checked / done / trash） |
    | 管理メモ | admin_memo |
    | 回答JSON | answers（[{questionId, value, files[]}]） |
    | 写真JSON | files（回答JSONの files を集めたもの。`collectFilesFromAnswers_`） |
    | 管理更新日時 | managed_at |

    **ゴミ箱は列ではなく対応状況。**`getResponses_` は `includeTrashed` が無いと status が "trash" の行を落とす。
    ここでも `trashed` は status で見る。

    アンケートごとのシート（`ensureSurveySheet_`。シート名 = アンケートのタイトル）は、
    見出しが「送信日時・回答ID・端末ID・お名前・メールアドレス・対応状況・管理メモ ＋ 設問の label」。
    値は回答JSONから写したものだが、シート上で直された可能性があるので `survey_sheet_values` に別に持つ
    （{見出し: 値}）。突き合わせに使い、正は answers。

    `getResponses_` が付け足す customerMemberNumber / customerNameKana は、端末IDかお名前で顧客プロフィールを
    引いた結果。**お名前で引く部分は写さない**（CLAUDE.md）。画面で要るなら CustomerProfile.client_ids で引く。
    """

    STATUS_CHOICES = [("new", "未対応"), ("checked", "確認済み"), ("done", "対応済み"), ("trash", "ゴミ箱")]

    # 回答一覧の何行目から来たか。取り込みの印。サーバーが自分で作った回答には無い
    sheet_row = models.IntegerField("回答一覧の行", null=True, blank=True, db_index=True)

    response_id = models.CharField("回答ID", max_length=128, unique=True)
    survey = models.ForeignKey(Survey, verbose_name="アンケート", null=True, blank=True,
                               on_delete=models.SET_NULL, related_name="responses")
    survey_key = models.CharField("アンケートID（GAS）", max_length=128, blank=True, db_index=True)
    survey_title = models.CharField("アンケート名", max_length=255, blank=True)

    submitted_at = models.DateTimeField("送信日時", null=True, blank=True, db_index=True)
    client_id = models.CharField("端末ID", max_length=128, blank=True, db_index=True)

    # 空のまま。**お名前から埋めない。**
    member_id = models.CharField("会員ID", max_length=32, blank=True, db_index=True)
    customer_name = models.CharField("お名前", max_length=255, blank=True, db_index=True)
    customer_email = models.CharField("メールアドレス", max_length=255, blank=True)

    status = models.CharField("対応状況", max_length=16, choices=STATUS_CHOICES, default="new", db_index=True)
    admin_memo = models.TextField("管理メモ", blank=True)

    answers = models.JSONField("回答", default=list, blank=True)
    files = models.JSONField("写真（回答一覧の写真JSON）", default=list, blank=True)
    # 読めなかったときは元の文字列を捨てずに残す
    answers_raw = models.TextField("回答（読めなかったもの）", blank=True)
    files_raw = models.TextField("写真（読めなかったもの）", blank=True)

    managed_at = models.DateTimeField("管理更新日時", null=True, blank=True)

    survey_sheet_row = models.IntegerField("アンケートのシートの行", null=True, blank=True)
    survey_sheet_values = models.JSONField("アンケートのシートの値", default=dict, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "回答"
        verbose_name_plural = "回答"
        ordering = ["-submitted_at", "-sheet_row"]
        indexes = [models.Index(fields=["survey_key", "-submitted_at"])]

    def __str__(self):
        return f"{self.customer_name} {self.survey_title} {self.submitted_at}"

    @property
    def trashed(self):
        return self.status == "trash"

    def answer_of(self, question_id):
        for a in self.answers or []:
            if isinstance(a, dict) and a.get("questionId") == question_id:
                return a.get("value")
        return None


class ResponsePhoto(models.Model):
    """回答の写真1枚。`savePhotoFiles_`（Code.gs 3260行〜）が Drive に置いた1ファイル。

    元は回答JSONの `answers[].files[]`（無ければ写真JSON）に
        {name, type, capturedAt, fileId, customerFolderName, customerFolderUrl, folderName, folderUrl,
         url, previewUrl, downloadUrl, thumbnailUrl}
    で入っている。**Drive のリンク（url / previewUrl / downloadUrl / thumbnailUrl / *FolderUrl）は持たない**
    （設計: 写真はサーバーの media へ移し、Drive の共有リンクは残さない）。
    ファイルIDだけ持ち、取り込み（ビジリスの回答を取り込む --写真）で media/bijiris/ へ入れて `url` を埋める。

    種別（kind）は Code.gs の判定（`measureTimingOf_` / TICKET_SURVEY_*_PHOTO_QUESTION_IDS）と同じ:
      - 計測時アンケート（q_measure_photos）は「計測のタイミング」の回答で決める。
        「初回計測」「モニター」を含めば before、「終了」を含めば after
      - 施術後アンケートの旧設問 *_monitor_photos* は before、*_ticket_end_photos* / q_ticket_end_photo_last は after
      - どれでもなければ other
    """

    KIND_CHOICES = [("before", "ビフォー"), ("after", "アフター"), ("other", "その他")]

    response = models.ForeignKey(Response, verbose_name="回答", on_delete=models.CASCADE, related_name="photos")
    question_id = models.CharField("設問ID", max_length=128, blank=True, db_index=True)
    drive_file_id = models.CharField("元の Drive ファイルID", max_length=128, db_index=True)
    name = models.CharField("ファイル名", max_length=255, blank=True)
    mime_type = models.CharField("種類", max_length=64, blank=True)
    # GAS が文字列のまま持っている（お客様の端末の時刻）。日時に直さずそのまま
    captured_at = models.CharField("撮影日時（端末の記録）", max_length=64, blank=True)
    customer_folder_name = models.CharField("Drive のお客様フォルダ名", max_length=255, blank=True)
    folder_name = models.CharField("Drive のフォルダ名", max_length=255, blank=True)
    # 取り込み（--写真）で media/bijiris/ に入れたときに埋まる。空なら「まだ Drive にしか無い」
    url = models.CharField("サーバーの URL", max_length=500, blank=True)
    order = models.IntegerField("並び", default=0)
    kind = models.CharField("種別", max_length=16, choices=KIND_CHOICES, default="other", db_index=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)

    class Meta:
        verbose_name = "回答の写真"
        verbose_name_plural = "回答の写真"
        ordering = ["response", "question_id", "order"]
        constraints = [
            models.UniqueConstraint(fields=["response", "question_id", "drive_file_id"], name="写真は回答と設問とファイルIDで一意"),
        ]

    def __str__(self):
        return f"{self.response_id} {self.question_id} {self.drive_file_id}"

    @property
    def stored(self):
        return bool(self.url)


class CustomerProfile(models.Model):
    """顧客1人ぶん。CUSTOMER_PROFILES_JSON（`normalizeCustomerProfileRecord_` 1903行〜）と
    CUSTOMER_MEMOS_JSON（`normalizeCustomerMemoRecord_` 1660行〜）を、**お名前を鍵**に1つにしたもの。

    | GAS | ここ |
    |---|---|
    | name（鍵） | name |
    | memberNumber | member_number（**GAS がまゆみの会員データを氏名で引いた結果**。`まゆみ会員_対応表_`。写すだけで member_id には入れない） |
    | nameKana / aliases / clientIds | name_kana / aliases / client_ids |
    | activeTicketCard {plan, sheetNumber, round} / activeTicketCardSource / lastTicketCardAcquiredAt | active_ticket_card / active_ticket_card_source / last_ticket_card_acquired_at |
    | measurementTargets {waist, hip, thighRight, thighLeft} | measurement_targets |
    | ticketStampAdjustment（-50〜50） | ticket_stamp_adjustment |
    | pushStatus {enabled, supported, permission, updatedAt} | push_status |
    | rewardRedemptions {"6": {handed, handedAt}, …}（特典の受け取り） | reward_redemptions |
    | adminManaged | admin_managed |
    | passcodeHash / passcodeSalt / passcodeUpdatedAt / passcodeSetupUntil | passcode_* |
    | updatedAt | updated_at |
    | メモ latestMemo / entries [{at(YYYY-MM-DD), memo}] | latest_memo / memo_entries |

    **パスコードのハッシュはビジリス GAS の TOKEN_SECRET を混ぜている**（`hashPasscode_`）。
    サーバーでは照合できない。持つのは「設定済みか」を見るためと、戻すときのため。
    """

    name = models.CharField("お名前", max_length=255, unique=True)
    # 空のまま。**お名前から埋めない。**
    member_id = models.CharField("会員ID", max_length=32, blank=True, db_index=True)
    member_number = models.CharField("会員番号（GAS の記録）", max_length=32, blank=True, db_index=True)
    name_kana = models.CharField("フリガナ", max_length=255, blank=True)
    aliases = models.JSONField("別名", default=list, blank=True)
    client_ids = models.JSONField("端末ID", default=list, blank=True)

    active_ticket_card = models.JSONField("回数券カード", null=True, blank=True)
    active_ticket_card_source = models.CharField("回数券カードの出どころ", max_length=16, blank=True)
    last_ticket_card_acquired_at = models.DateTimeField("回数券カードの取得日時", null=True, blank=True)
    measurement_targets = models.JSONField("計測の目標", null=True, blank=True)
    ticket_stamp_adjustment = models.IntegerField("回数券スタンプの手当て", default=0)
    push_status = models.JSONField("通知の状態", null=True, blank=True)
    reward_redemptions = models.JSONField("特典の受け取り", null=True, blank=True)
    admin_managed = models.BooleanField("受付が作った", default=False)

    passcode_hash = models.CharField("パスコードのハッシュ", max_length=128, blank=True)
    passcode_salt = models.CharField("パスコードのソルト", max_length=128, blank=True)
    passcode_updated_at = models.CharField("パスコードの更新日時", max_length=64, blank=True)
    passcode_setup_until = models.CharField("再設定の許可期限", max_length=64, blank=True)

    latest_memo = models.TextField("最新のメモ", blank=True)
    memo_entries = models.JSONField("メモの履歴", default=list, blank=True)

    updated_at = models.DateTimeField("更新日時", null=True, blank=True)

    imported_at = models.DateTimeField("取り込み日時", auto_now_add=True)
    changed_at = models.DateTimeField("変更日時", auto_now=True)

    class Meta:
        verbose_name = "顧客"
        verbose_name_plural = "顧客"
        ordering = ["name"]

    def __str__(self):
        return self.name

    @property
    def has_passcode(self):
        return bool(self.passcode_hash and self.passcode_salt)

    @property
    def redeemed_thresholds(self):
        """受け取り済みの特典の節目（回数）。"""
        return sorted(int(k) for k in (self.reward_redemptions or {}).keys() if str(k).isdigit())
