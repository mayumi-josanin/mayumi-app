"""ビジリスの設定（ADMIN_PREFERENCES_JSON）。records.AppSetting の鍵 `bijiris_preferences` に中身そのまま持つ。

既定値と正規化は、ビジリス GAS の `getPreferences_`（Code.gs 1058行〜）と同じ。
**片方だけ直すと、切り替えの前後で設定の見え方が変わる。**直すときは両方。

GAS と違うところ:
  - notificationEmail の既定は GAS では `getOwnerEmail_()`（スクリプトの持ち主）。サーバーには持ち主が無いので空
  - twoFactorEnabled は GAS でも常に false（設定画面は作らないので、そのまま）
"""

import re

# ---- 既定の文言（DEFAULT_*_ 関数の写し） ----
既定の通知件名 = "【まゆみ助産院】新しいアンケート回答"
既定の通知本文 = "\n".join([
    "新しいアンケート回答が届きました。",
    "",
    "お名前: {{customerName}}",
    "アンケート: {{surveyTitle}}",
    "送信日時: {{submittedAt}}",
    "回答ID: {{responseId}}",
])
既定の利用目的文 = "ご回答内容と添付写真は、まゆみ助産院のアンケート管理と施術サポートのために利用します。"
既定の同意文 = "回答内容と添付写真の利用に同意します。"
既定の復旧メモ = "障害時は Apps Script の実行ログ、スプレッドシート、Google ドライブの保存フォルダを確認してください。"
既定の一般カテゴリ = ["豆知識", "アドバイス", "セルフケア", "お知らせ", "よくある質問"]
既定のお悩みの根 = "お悩み"
既定のお悩みの道 = [
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > 産後の骨盤底筋のゆるみ感が気になる",
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > くしゃみ、咳、大笑いでヒヤッとする",
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > 骨盤まわりを土台からケアしたい",
    "女性 > トイレまわりのお悩み > 尿漏れが気になる",
    "女性 > トイレまわりのお悩み > 頻尿が気になる",
    "女性 > トイレまわりのお悩み > 急な尿意が気になる",
    "女性 > トイレまわりのお悩み > 便秘がち",
    "女性 > トイレまわりのお悩み > お通じのリズムが気になる",
    "女性 > 体型・見た目のお悩み > 産後のぽっこりお腹が気になる",
    "女性 > 体型・見た目のお悩み > 年齢とともに体型の変化が気になる",
    "女性 > 体型・見た目のお悩み > 下半身太りが気になる",
    "女性 > 体型・見た目のお悩み > ヒップの下垂が気になる",
    "女性 > 姿勢・日常動作のお悩み > 産後に姿勢が崩れやすくなった",
    "女性 > 姿勢・日常動作のお悩み > 抱っこや家事で下腹や骨盤まわりが気になる",
    "女性 > 姿勢・日常動作のお悩み > 姿勢を整えたい",
    "女性 > デリケートゾーンまわりのお悩み > デリケートゾーンのケアを意識したい",
    "女性 > デリケートゾーンまわりのお悩み > 膣トレを始めてみたい",
    "女性 > 冷え・巡りのお悩み > 冷えやすさが気になる",
    "男性 > トイレまわりのお悩み > 頻尿が気になる",
    "男性 > トイレまわりのお悩み > ちょい漏れが気になる",
    "男性 > トイレまわりのお悩み > 急な尿意で不安がある",
    "男性 > トイレまわりのお悩み > トイレ悩みをケアしたい",
    "男性 > デリケートなお悩み > EDケアを意識したい",
    "男性 > デリケートなお悩み > デリケートなお悩みを人知れずケアしたい",
    "男性 > 姿勢・骨盤まわりのお悩み > 長時間の座り仕事で骨盤まわりが気になる",
    "男性 > 姿勢・骨盤まわりのお悩み > 猫背や前かがみ姿勢が気になる",
    "男性 > 姿勢・骨盤まわりのお悩み > 腰まわりの違和感が気になる",
    "男性 > 下半身・体型のお悩み > 下半身の筋力低下が気になる",
    "男性 > 下半身・体型のお悩み > ヒップラインの崩れが気になる",
    "男性 > 下半身・体型のお悩み > むくみや冷えが気になる",
    "男性 > 下半身・体型のお悩み > 運動不足が気になる",
    "男性 > 下半身・体型のお悩み > 筋トレが続かない",
]
ガチャの景品の鍵 = ("A", "B", "C", "D")

# 回数券分析の AI への指示文の既定（TICKET_SURVEY_DEFAULT_PROMPT）。TICKET_SURVEY_PROMPT が空のときはこれ
既定の分析指示文 = "\n".join([
    "あなたはまゆみ助産院の EMS トレーニング「ビジリス」の施術者です。",
    "同一のお客様の「モニター時（ビフォー）」と「回数券終了時（アフター）」の全身写真を比較し、",
    "身体の変化をお客様にお伝えするための分析文を日本語で作成してください。",
    "",
    "【観察してほしい観点】",
    "1. 姿勢（骨盤の前後傾・反り腰・猫背・肩の高さ・頭の位置）",
    "2. お腹まわり（下腹のふくらみ、ウエストのくびれ）",
    "3. ヒップの位置と丸み、太もものライン",
    "4. 全体のシルエットと立ち姿の安定感",
    "",
    "【出力形式】",
    "■ 変化のポイント（3つ、それぞれ2〜3文）",
    "■ 特に良くなった点（1〜2文）",
    "■ これから伸ばせる点とおすすめの続け方（2〜3文）",
    "",
    "【注意】",
    "・医学的な診断や断定は避け、見た目の変化の範囲で書いてください。",
    "・お客様ご本人が読んで前向きになれる、やさしく丁寧な敬体で書いてください。",
    "・写真から読み取れないことは推測で断定せず、書かないでください。",
])

鍵 = "bijiris_preferences"
回数券メタの鍵 = "bijiris_ticket_meta"
回数券指示文の鍵 = "bijiris_ticket_prompt"


def 文字(値) -> str:
    return "" if 値 is None else str(値).strip()


def _数(値):
    try:
        n = float(値)
    except (TypeError, ValueError):
        return None
    if n != n or n in (float("inf"), float("-inf")):
        return None
    return n


def 文字の一覧(値, 既定):
    出 = [文字(v) for v in (値 if isinstance(値, list) else [])]
    出 = [v for v in 出 if v]
    return 出 if 出 else list(既定)


def 月の鍵(値) -> str:
    m = re.match(r"^(\d{4})[-/年](\d{1,2})", 文字(値))
    if not m:
        return ""
    y, mo = int(m.group(1)), int(m.group(2))
    if not 1 <= mo <= 12:
        return ""
    return f"{y:04d}-{mo:02d}"


def 確率(値, 既定):
    n = _数(値)
    if n is None:
        return 既定
    return max(0.0, min(100.0, round(n * 10) / 10))


def 利用目的文を整える(値) -> str:
    s = re.sub(r"保存先は Google スプレッドシートおよび Google ドライブです。?", "", 文字(値))
    return re.sub(r"\s+", " ", s).strip()


def ガチャ設定を整える(値, 今月: str = "") -> dict:
    """normalizeGachaPrizeConfig_ と同じ。月ごとに A〜D の景品と確率。空なら今月ぶんを1つ作る。"""
    見た = set()
    月々 = []
    for e in (値.get("monthlyPrizes") if isinstance(値, dict) and isinstance(値.get("monthlyPrizes"), list) else []):
        月 = 月の鍵(e.get("month") if isinstance(e, dict) else None)
        if not 月 or 月 in 見た:
            continue
        見た.add(月)
        景品 = {}
        for k in ガチャの景品の鍵:
            p = (e.get("prizes") or {}).get(k) if isinstance(e, dict) and isinstance(e.get("prizes"), dict) else None
            p = p if isinstance(p, dict) else {}
            景品[k] = {"content": 文字(p.get("content")), "probability": 確率(p.get("probability"), 0)}
        月々.append({"month": 月, "prizes": 景品})
    月々.sort(key=lambda x: x["month"])
    if not 月々:
        from django.utils import timezone

        月 = 今月 or 月の鍵(timezone.now().isoformat()) or "2026-04"
        月々 = [{"month": 月, "prizes": {k: {"content": "", "probability": 25} for k in ガチャの景品の鍵}}]
    return {"monthlyPrizes": 月々}


def 節目特典を整える(値) -> dict:
    """normalizeMilestoneRewardConfig_ と同じ。回数（threshold）ごとの特典。最大20。"""
    有効 = not (isinstance(値, dict) and 値.get("enabled") is False)
    節目 = {}
    for e in (値.get("milestones") if isinstance(値, dict) and isinstance(値.get("milestones"), list) else []):
        if not isinstance(e, dict):
            continue
        n = _数(e.get("threshold"))
        reward = 文字(e.get("reward"))
        if n is None or int(n) <= 0 or not reward:
            continue
        t = int(n)
        節目[t] = {"threshold": t, "reward": reward, "description": 文字(e.get("description"))}
    return {"enabled": 有効, "milestones": [節目[k] for k in sorted(節目)][:20]}


def 分類を整える(値) -> dict:
    値 = 値 if isinstance(値, dict) else {}
    return {
        "generalCategories": 文字の一覧(値.get("generalCategories"), 既定の一般カテゴリ),
        "concernRootLabel": 文字(値.get("concernRootLabel")) or 既定のお悩みの根,
        "concernPaths": 文字の一覧(値.get("concernPaths"), 既定のお悩みの道),
    }


def 控えの時刻を整える(値) -> int:
    n = _数(値)
    return int(n) if n is not None and 0 <= n <= 23 else 3


def 保管日数を整える(値) -> int:
    n = _数(値)
    return int(n // 1) if n is not None and n >= 0 else 365


def 正規化する(stored) -> dict:
    """getPreferences_ と同じ既定値で、欠けを埋めた設定を返す。"""
    s = stored if isinstance(stored, dict) else {}
    return {
        "notificationEnabled": s.get("notificationEnabled") is not False,
        "notificationEmail": 文字(s.get("notificationEmail")).lower(),
        "notificationSubject": 文字(s.get("notificationSubject")) or 既定の通知件名,
        "notificationBody": 文字(s.get("notificationBody")) or 既定の通知本文,
        "dataPolicyText": 利用目的文を整える(s.get("dataPolicyText")) or 既定の利用目的文,
        "requireConsent": s.get("requireConsent") is not False,
        "consentText": 文字(s.get("consentText")) or 既定の同意文,
        "autoBackupEnabled": s.get("autoBackupEnabled") is not False,
        "backupHour": 控えの時刻を整える(s.get("backupHour")),
        "retentionDays": 保管日数を整える(s.get("retentionDays")),
        "recoveryMemo": 文字(s.get("recoveryMemo")) or 既定の復旧メモ,
        "twoFactorEnabled": False,
        "bijirisCategoryConfig": 分類を整える(s.get("bijirisCategoryConfig")),
        "gachaPrizeConfig": ガチャ設定を整える(s.get("gachaPrizeConfig")),
        "milestoneRewardConfig": 節目特典を整える(s.get("milestoneRewardConfig")),
        "campaignStampEnabled": s.get("campaignStampEnabled") is not False,
    }


def 既定値() -> dict:
    return 正規化する({})


def 読む() -> dict:
    """AppSetting から読み、欠けを既定で埋めて返す（無ければ既定）。"""
    from apps.records.models import AppSetting

    s = AppSetting.objects.filter(pk=鍵).first()
    return 正規化する(s.value if s else {})


def 分析指示文を読む() -> str:
    from apps.records.models import AppSetting

    s = AppSetting.objects.filter(pk=回数券指示文の鍵).first()
    return (文字((s.value or {}).get("prompt")) if s else "") or 既定の分析指示文
