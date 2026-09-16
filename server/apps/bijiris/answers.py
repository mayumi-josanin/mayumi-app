"""回答の読み方（旧ビジリス管理アプリ admin-app/app.js の共通関数の写し）。

集計・アンケート管理・回答管理の3画面で同じ数え方をするために、ここに1つだけ置く。
app.js の関数名との対応:

| app.js | ここ |
|---|---|
| STATUS_LABELS / normalizeStatus | 対応状況の名 / 対応状況を整える |
| TICKET_INFO_QUESTION_IDS / getResponseTicketInfo | 回数券の設問ID / 回数券情報 |
| SESSION_CONCERN_CATEGORIES / getConcernAnswerGroups | お悩みの分類 / お悩みの群 |
| SESSION_LIFE_CHANGE_CATEGORIES / getCategorizedAnswerGroups | 変化の分類 / 変化の群 |
| getAnswerValues / getCheckboxAnswerValues | 回答の値一覧 |
| getDisplayAnswers | 表示用の回答 |
| collectPhotosFromResponse | 写真の一覧（ResponsePhoto から。Drive のリンクは持たない） |
| getCustomerNameWithMember | 顧客の控え.会員番号付きの名前 |
| formatAnswerForCsv | CSV用の回答文 |
| getResponseSearchHaystack | 検索用の文 |
| formatDate | 日時の文字 |

**お名前で会員を探さない**（CLAUDE.md）。会員番号は CustomerProfile に写してある `member_number` を
端末ID（client_ids）で引き、無ければお名前で引いて**表示に添えるだけ**。会員の表は見ない。
"""

from django.utils import timezone

対応状況の名 = {"new": "未対応", "checked": "確認済み", "done": "対応済み", "trash": "ゴミ箱"}
対応状況の印 = {"new": "badge-orange", "checked": "badge-blue", "done": "badge-green", "trash": "badge-gray"}

# 旧アプリの TICKET_INFO_QUESTION_IDS。施術後アンケート（新旧）の回数券の設問
回数券の設問ID = {
    "回数券": ["q_bijiris_session_ticket_plan", "q_ticket_end_ticket_size"],
    "何枚目": ["q_bijiris_session_ticket_sheet", "q_ticket_end_ticket_sheet"],
    "何回目": ["q_bijiris_session_ticket_round", "q_ticket_end_ticket_round"],
}
お悩みの設問ID = "q_bijiris_session_concern"
変化の設問ID = "q_bijiris_session_life_changes"

# 回答管理の絞り込みの選択肢（CUSTOMER_TICKET_*_OPTIONS）
回数券の種類 = ["6回券", "10回券"]
回数券の枚目 = [f"{i}枚目" for i in range(1, 21)]
回数券の回目 = [f"{i}回目" for i in range(1, 11)]

設問の種類 = ["text", "textarea", "choice", "checkbox", "rating", "photo"]
設問の種類の名 = {"text": "テキスト", "textarea": "長文", "choice": "単一選択", "checkbox": "複数選択", "rating": "5段階評価", "photo": "写真"}


def 対応状況を整える(status) -> str:
    return status if status in 対応状況の名 else "new"


def 文字(値) -> str:
    if 値 is None:
        return ""
    if isinstance(値, list):
        return ", ".join(文字(v) for v in 値 if 文字(v))
    return str(値).strip()


def 日時の文字(dt) -> str:
    """formatDate: 2026/08/10 10:05（日本時間）。無ければ「-」。"""
    if not dt:
        return "-"
    return timezone.localtime(dt).strftime("%Y/%m/%d %H:%M")


def 回答の値一覧(answer) -> list[str]:
    """複数選択の回答は「a, b, c」の1文字列。カンマで割る（getAnswerValues）。"""
    値 = (answer or {}).get("value") if isinstance(answer, dict) else answer
    if isinstance(値, list):
        return [文字(v) for v in 値 if 文字(v)]
    return [v.strip() for v in 文字(値).split(",") if v.strip()]


def 回答の表(response) -> dict:
    """questionId → 回答（dict）。"""
    表 = {}
    for a in response.answers or []:
        if isinstance(a, dict) and a.get("questionId"):
            表[a["questionId"]] = a
    return 表


def _最初の値(表, 設問IDたち) -> str:
    for qid in 設問IDたち:
        v = 文字((表.get(qid) or {}).get("value"))
        if v:
            return v
    return ""


def 回数券情報(response) -> list[tuple[str, str]]:
    """[("回数券","10回券"),("何枚目","1枚目"),("何回目","3回目")]。値が無い項目は入れない。"""
    表 = 回答の表(response)
    出 = []
    for label, ids in 回数券の設問ID.items():
        v = _最初の値(表, ids)
        if v:
            出.append((label, v))
    return 出


def 分類の群(answer, 分類, その他の名) -> list[dict]:
    """選んだ選択肢を分類ごとにまとめる。分類に無い選択肢は「その他」の群（id: extras）。"""
    選んだ = 回答の値一覧(answer)
    集合 = set(選んだ)
    群 = []
    for c in 分類:
        当たり = [o for o in c["options"] if o in 集合]
        if 当たり:
            群.append({"id": c["id"], "label": c["label"], "options": 当たり, "count": len(当たり)})
    既知 = {o for c in 分類 for o in c["options"]}
    余り = [o for o in 選んだ if o not in 既知]
    if 余り and その他の名:
        群.append({"id": "extras", "label": その他の名, "options": 余り, "count": len(余り)})
    return 群


def お悩みの群(answer):
    return 分類の群(answer, お悩みの分類, "【その他】")


def 変化の群(answer):
    # getCategorizedAnswerGroups は分類に無いものを捨てる（その他の群を作らない）
    return 分類の群(answer, 変化の分類, "")


def お悩みの分類ID(response) -> list[str]:
    return [g["id"] for g in お悩みの群(回答の表(response).get(お悩みの設問ID))]


def お悩みの分類名(response) -> list[str]:
    return [g["label"] for g in お悩みの群(回答の表(response).get(お悩みの設問ID))]


def 表示用の回答(response, survey) -> list[dict]:
    """アンケートの設問の並びで回答を並べ、設問に無い回答は後ろに足す（getDisplayAnswers）。"""
    answers = [a for a in (response.answers or []) if isinstance(a, dict)]

    def 整える(a, q=None):
        q = q or {}
        return dict(a, label=a.get("label") or q.get("label") or a.get("questionId") or "質問",
                    type=a.get("type") or q.get("type") or "text", value=a.get("value", ""))

    if not survey:
        return [整える(a) for a in answers]
    表 = {a.get("questionId"): a for a in answers}
    出 = []
    for q in survey.questions or []:
        if not isinstance(q, dict):
            continue
        a = 表.get(q.get("id"))
        if a:
            出.append(整える(a, q))
        else:
            出.append({"questionId": q.get("id"), "label": q.get("label") or "", "type": q.get("type") or "text", "value": ""})
    設問ID = {q.get("id") for q in survey.questions or [] if isinstance(q, dict)}
    出.extend(整える(a) for a in answers if a.get("questionId") not in 設問ID)
    return 出


def 設問の表(survey) -> dict:
    return {q.get("id"): q for q in ((survey.questions if survey else None) or []) if isinstance(q, dict)}


def 写真の一覧(response) -> list:
    """回答の写真（ResponsePhoto）。prefetch してあれば通信しない。"""
    return list(response.photos.all())


def 写真を設問ごとに(response) -> dict:
    出 = {}
    for p in 写真の一覧(response):
        出.setdefault(p.question_id, []).append(p)
    return 出


class 顧客の控え:
    """CustomerProfile を端末IDとお名前で引ける形にしたもの。1画面で1回だけ作る。"""

    def __init__(self, profiles):
        self.端末で = {}
        self.名前で = {}
        for p in profiles:
            self.名前で[p.name] = p
            for cid in p.client_ids or []:
                if cid:
                    self.端末で.setdefault(str(cid), p)

    def 引く(self, response):
        if response.client_id and response.client_id in self.端末で:
            return self.端末で[response.client_id]
        return self.名前で.get((response.customer_name or "").strip())

    def 会員番号(self, response) -> str:
        p = self.引く(response)
        return (p.member_number or "").strip().upper() if p else ""

    def フリガナ(self, response) -> str:
        p = self.引く(response)
        return (p.name_kana or "") if p else ""

    def 会員番号付きの名前(self, response) -> str:
        番号 = self.会員番号(response)
        return f"{番号} / {response.customer_name}" if 番号 else (response.customer_name or "")


def CSV用の回答文(answer, 写真たち) -> str:
    label = answer.get("label") or ""
    if 写真たち:
        return f"{label}: " + ", ".join((p.url or p.name or "未取り込み") for p in 写真たち)
    if answer.get("questionId") == お悩みの設問ID:
        return f"{label}: " + " | ".join(f"{g['label']} " + " / ".join(g["options"]) for g in お悩みの群(answer))
    return f"{label}: {文字(answer.get('value'))}"


def 回答文の一覧(response) -> list[str]:
    写真 = 写真を設問ごとに(response)
    return [CSV用の回答文(a, 写真.get(a.get("questionId"), [])) for a in (response.answers or []) if isinstance(a, dict)]


def 検索用の文(response, 控え: 顧客の控え) -> str:
    return " ".join(x for x in [
        response.customer_name or "",
        控え.会員番号(response),
        控え.フリガナ(response),
        response.survey_title or "",
        response.admin_memo or "",
        " ".join(f"{k} {v}" for k, v in 回数券情報(response)),
        " ".join(お悩みの分類名(response)),
        " ".join(回答文の一覧(response)),
    ] if x).lower()


def ゆるい文字(値) -> str:
    """normalizeLooseSearchText: 空白を全部抜いて小文字に。"""
    return "".join(文字(値).split()).lower()


# app.js の SESSION_CONCERN_CATEGORIES（施術後アンケート「お悩み」の分類。7分類×10項目）
お悩みの分類 = [
    {"id": "toilet", "label": "【トイレ・デリケートなお悩み】", "options": [
        "咳やくしゃみ、大笑いをした時に少し気になることがある",
        "ジャンプや運動、重いものを持った時に気になることがある",
        "急にトイレに行きたくなり、間に合うか不安になることがある",
        "以前よりトイレが近くなった気がする",
        "夜中にトイレで目が覚めることがある",
        "外出先でトイレの場所が気になりやすい",
        "トイレのあとも、すっきりしない感じがある",
        "尿の出が弱い、出にくいと感じることがある",
        "トイレに時間がかかることがある",
        "ナプキンやパッドが手放せず不安を感じることがある",
    ]},
    {"id": "belly", "label": "【お腹まわり・便通のお悩み】", "options": [
        "便秘が気になる",
        "すっきり出にくいと感じることがある",
        "お腹に力を入れにくい感じがある",
        "下腹が張りやすい",
        "下腹ぽっこりが気になる",
        "お腹まわりの支えが弱くなった気がする",
        "インナーマッスルの衰えが気になる",
        "お腹まわりをすっきり整えたい",
        "お腹の奥に力が入りにくい感じがある",
        "体の内側から支えられていない感じがある",
    ]},
    {"id": "pelvic", "label": "【骨盤まわり・内側の筋力のお悩み】", "options": [
        "骨盤まわりが不安定に感じる",
        "骨盤底筋をうまく使えていない気がする",
        "締める感覚がわかりにくい",
        "自分では鍛えにくい部分をケアしたい",
        "体の内側の支える力が弱くなった気がする",
        "体幹の弱さが気になる",
        "出産後から骨盤まわりの変化が気になる",
        "年齢とともに筋力の低下を感じる",
        "将来のために早めにケアしておきたい",
        "骨盤の底から支える感覚を取り戻したい",
    ]},
    {"id": "posture", "label": "【姿勢・体型のお悩み】", "options": [
        "姿勢の崩れが気になる",
        "猫背が気になる",
        "反り腰が気になる",
        "立ち姿をきれいに見せたい",
        "歩き方や姿勢を整えたい",
        "下腹ぽっこりが気になる",
        "ヒップラインの変化が気になる",
        "ヒップアップしたい",
        "体のラインをすっきり整えたい",
        "無理なく体の土台から整えたい",
    ]},
    {"id": "lower-body", "label": "【腰まわり・下半身のお悩み】", "options": [
        "腰まわりに負担を感じやすい",
        "股関節まわりが硬く感じる",
        "お尻の筋肉をうまく使えていない気がする",
        "太ももばかり疲れやすい",
        "長時間立っているとつらい",
        "歩くと疲れやすい",
        "階段の上り下りが気になる",
        "下半身の筋力低下が気になる",
        "下半身を安定させたい",
        "お尻や骨盤まわりをしっかり使えるようになりたい",
    ]},
    {"id": "postpartum-aging", "label": "【産後・年齢による変化】", "options": [
        "出産後から体の変化が気になっている",
        "出産後から骨盤まわりが不安定に感じる",
        "出産後、お腹やお尻まわりが戻りにくいと感じる",
        "以前より体を支える力が弱くなった気がする",
        "年齢とともに変化を感じるようになった",
        "更年期以降、トイレや骨盤まわりの悩みが増えた",
        "今は大きな悩みはないが、予防として始めたい",
        "将来のために骨盤底筋ケアを取り入れたい",
        "出産後、咳や抱っこで気になることが増えた",
        "これから先の体の変化に備えて整えておきたい",
    ]},
    {"id": "daily-life", "label": "【日常生活で気になること】", "options": [
        "外出や旅行の時に少し不安がある",
        "長時間の移動が気になる",
        "会議や授業中にトイレが気になることがある",
        "運動や趣味を思いきり楽しみにくい",
        "ジャンプやランニングを控えることがある",
        "重い荷物を持つ時に不安がある",
        "お子さまの抱っこなどで気になることがある",
        "夜ぐっすり眠りたい",
        "日常のちょっとした動作に不安がある",
        "トイレを気にせず過ごせる時間を増やしたい",
    ]},
]

# app.js の SESSION_LIFE_CHANGE_CATEGORIES（施術後アンケート「生活の変化」の分類）
変化の分類 = [
    {"id": "toilet-change", "label": "【トイレまわりの変化】", "options": [
        "咳やくしゃみをした時の不安が以前より減った",
        "急な尿意を気にする場面が減った",
        "外出時にトイレの場所を気にしすぎなくなった",
        "夜中にトイレで起きる回数が減った",
    ]},
    {"id": "core-change", "label": "【お腹・骨盤まわりの変化】", "options": [
        "お腹の奥に力が入りやすくなった",
        "骨盤まわりが安定した感じがある",
    ]},
    {"id": "posture-change", "label": "【姿勢・見た目の変化】", "options": [
        "姿勢を意識しやすくなった",
        "下腹まわりがすっきりした感じがある",
    ]},
    {"id": "movement-change", "label": "【動きやすさの変化】", "options": [
        "歩く・立つ・動くことが以前より楽になった",
    ]},
    {"id": "other-change", "label": "【その他】", "options": [
        "その他（自由記述）",
    ]},
]
