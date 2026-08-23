"""GAS と同じ形で受ける**書き込み**の窓口。

読み取り（`views.py`）と分けてある。**性質が違うから。**

読み取りは、間違えても表示が変わるだけで、GASと突き合わせれば分かる。
書き込みは**お客様の記録が変わる。**間違えると元に戻せないことがあり、
同じ操作を2回受ければ2回書き込まれる。

## ここで守ること

1. **2回届いても壊れない**（通信は失敗して再送される）
2. **同時に届いても壊れない**（行ロックを使う）
3. **GASと同じ答えを返す**（アプリは返ってきた中身で画面を描く）
4. **書く前に必ず会員を確かめる**（無い会員に書かない）
"""

import json
import random
import re
import unicodedata

from django.contrib.auth.hashers import make_password
from django.db import transaction
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_datetime

from apps.members.models import Member

# GAS の MAX_DEVICE_SESSIONS と同じ。端末の記録は8件まで。
端末の上限 = 8


def _日時の文字(d):
    """GAS の formatDateTime_ と同じ形。"""
    return d.astimezone().strftime("%Y-%m-%dT%H:%M:%S%z").replace("+0900", "+09:00")


def _端末を整える(生, いま今の端末ID=""):
    """GAS の normalizeDeviceSessionList_ と同じ。

    **deviceId が無いものは捨てる。**最後に使った順に並べ、8件までにする。
    `=== true` のときだけ真にするのも同じ（"false" という文字列を真にしない）。
    """
    if not isinstance(生, list):
        return []
    出 = []
    for d in 生:
        if not isinstance(d, dict):
            continue
        端末ID = str(d.get("deviceId") or "").strip()
        if not 端末ID:
            continue
        出.append({
            "deviceId": 端末ID,
            "label": str(d.get("label") or "").strip(),
            "platform": str(d.get("platform") or "").strip(),
            "appVersion": str(d.get("appVersion") or "").strip(),
            "lastSeenAt": str(d.get("lastSeenAt") or "").strip() or _日時の文字(timezone.now()),
            "passcodeEnabled": d.get("passcodeEnabled") is True,
            "pushEnabled": d.get("pushEnabled") is True,
            "current": d.get("current") is True,
        })
    出.sort(key=lambda x: x["lastSeenAt"] or "", reverse=True)
    return 出[:端末の上限]


def 端末をそろえる(data):
    """GAS の handleSyncUserDeviceSession()（7660行）と同じ。

    アプリを開くたびに呼ばれる。**93日で7,788回**と、いちばん多い書き込み。

    やること: その端末の記録を**入れ替える**（無ければ足す）。
    同じ端末IDが2つ並ばないよう、先に消してから先頭に足す。

    **2回届いても安全。**同じ端末IDなら上書きになるだけで、増えない。
    ただし `lastSeenAt` は書き換わる。GASも同じ。
    """
    会員ID = str((data or {}).get("memberId") or "").strip()
    端末ID = str((data or {}).get("deviceId") or "").strip()

    # GAS と同じ文言。アプリはこの文言を見ていないが、記録には残る。
    if not 会員ID or not 端末ID:
        return {"status": "error", "message": "会員IDまたは端末IDが不足しています"}

    # **行ロックを取る。**同じ会員に同時に2つ届くと、
    # 後から書いたほうが前の端末を消してしまう。
    with transaction.atomic():
        m = Member.objects.select_for_update().filter(member_id=会員ID).first()
        if not m:
            return {"status": "error", "message": "会員情報が見つかりません"}

        いま = _日時の文字(timezone.now())

        # その端末を一度取り除いてから、先頭に足す（GASと同じ順序）
        残り = [
            d for d in _端末を整える(m.device_sessions)
            if d["deviceId"] != 端末ID
        ]
        次 = [{
            "deviceId": 端末ID,
            "label": str((data or {}).get("label") or "").strip(),
            "platform": str((data or {}).get("platform") or "").strip(),
            "appVersion": str((data or {}).get("appVersion") or "").strip(),
            "lastSeenAt": いま,
            "passcodeEnabled": (data or {}).get("passcodeEnabled") is True,
            "pushEnabled": (data or {}).get("pushEnabled") is True,
            "current": True,
        }] + 残り

        整えた = _端末を整える(次)
        # **いま使っている端末だけ current を真にする。**GASと同じ。
        for d in 整えた:
            d["current"] = d["deviceId"] == 端末ID

        m.device_sessions = 整えた
        m.save(update_fields=["device_sessions", "changed_at"])

    return {"status": "ok", "devices": 整えた}


def 端末を外す(data):
    """GAS の handleRemoveUserDeviceSession()（7695行）と同じ。

    指定された端末の記録だけを消す。**2回届いても安全**（もう無いだけ）。
    """
    会員ID = str((data or {}).get("memberId") or "").strip()
    端末ID = str((data or {}).get("deviceId") or "").strip()
    if not 会員ID or not 端末ID:
        return {"status": "error", "message": "会員IDまたは端末IDが不足しています"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(member_id=会員ID).first()
        if not m:
            return {"status": "error", "message": "会員情報が見つかりません"}
        残り = [d for d in _端末を整える(m.device_sessions) if d["deviceId"] != 端末ID]
        m.device_sessions = 残り
        m.save(update_fields=["device_sessions", "changed_at"])

    return {"status": "ok", "devices": 残り}



def _数(値, 既定, 最小=None, 最大=None):
    """数に直す。**読めなければ既定値を返す（0にしない）。**

    GAS は `Number(x) || 0` と書いていて、これは
    **null・空文字・文字列が来ると 0 になる。**
    スタンプでこれが起きると、お客様の記録が0に戻る。

    CLAUDE.md の「記録を読む通信で書き込む」を禁ずる項目は、
    まさにこの形の不具合（受付で入れたスタンプが0に戻った）から来ている。
    **ここでは 0 に落とさず、いまの値を残す。**
    """
    if 値 is None:
        return 既定
    try:
        n = int(値)
    except (TypeError, ValueError):
        try:
            n = int(float(値))
        except (TypeError, ValueError):
            return 既定
    if 最小 is not None:
        n = max(最小, n)
    if 最大 is not None:
        n = min(最大, n)
    return n


def _一覧(値, 既定):
    """配列に直す。読めなければ既定値。**空配列にしない。**"""
    if 値 is None:
        return 既定
    if isinstance(値, list):
        return 値
    if isinstance(値, str):
        try:
            v = json.loads(値)
            return v if isinstance(v, list) else 既定
        except ValueError:
            return 既定
    return 既定


def 特典をそろえる(data):
    """GAS の handleSyncUserRewardStatus()（7509行）と同じ。

    アプリが持っているスタンプ・特典の状態を送ってきて、記録に反映する。
    **93日で791回**呼ばれる。

    ## ここがいちばん危ない

    アプリから送られた値で**上書きする**作り。
    送られてこなかった項目は、**いまの値を残す**（GAS も同じ）。

        data.stampCount が undefined  → いまの値を使う
        data.stampCount が 5          → 5 にする

    **問題は「読めない値」が来たとき。**
    GAS は `Number(x) || 0` なので、`null` や文字列が来ると **0 になる。**
    お客様のスタンプが0に戻る。

    **ここでは 0 に落とさず、いまの値を残す。**
    GAS と違う動きになるが、**こちらのほうが安全side。**
    「スタンプを0にする」は、管理アプリから明示的に行うべきこと。

    ## 会員が無いときは作らない

    GAS は `ensureUserRowFromActivity_` で**新しい行を足す**。
    ここでは足さない。**足すのは updateUser の仕事**にする。
    無い会員に特典だけ書くと、名前も電話番号も無い行ができる。
    （2026-08 に「入れないお客様」が21名いた原因が、まさにこれ）
    """
    会員ID = str((data or {}).get("memberId") or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが指定されていません"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(member_id=会員ID).first()
        if not m:
            # **GASと違って行を作らない。**理由は上のとおり。
            return {"status": "error", "message": "会員情報が見つかりません"}

        d = data or {}
        # 送られてこなかった項目は、いまの値を残す。
        # 読めない値が来たときも、いまの値を残す（0にしない）。
        次 = {
            "stampCount": _数(d.get("stampCount"), m.stamp_count or 0, 0, 10),
            "stampCardNum": _数(d.get("stampCardNum"), m.stamp_card_number or 1, 1),
            "rewards": _一覧(d.get("rewards"), m.reward_history or []),
            "stampHistory": _一覧(d.get("stampHistory"), m.stamp_history or []),
        }

        m.stamp_count = 次["stampCount"]
        m.stamp_card_number = 次["stampCardNum"]
        m.reward_history = 次["rewards"]
        m.stamp_history = 次["stampHistory"]

        # 最終スタンプ日時。送られてきたものがあれば使う。
        生 = d.get("lastStampAt") or d.get("lastStampDate")
        if 生:
            from django.utils.dateparse import parse_datetime, parse_date

            t = parse_datetime(str(生))
            if t is None:
                日 = parse_date(str(生)[:10])
                if 日:
                    t = timezone.make_aware(
                        timezone.datetime.combine(日, timezone.datetime.min.time()))
            if t is not None:
                if timezone.is_naive(t):
                    t = timezone.make_aware(t)
                m.last_stamp_at = t

        m.save(update_fields=[
            "stamp_count", "stamp_card_number", "reward_history",
            "stamp_history", "last_stamp_at", "changed_at",
        ])

        返す = {
            "stampCount": m.stamp_count,
            "stampCardNum": m.stamp_card_number,
            "rewards": m.reward_history or [],
            "stampHistory": m.stamp_history or [],
            "lastStampDate": m.last_stamp_at.astimezone().strftime("%Y-%m-%d") if m.last_stamp_at else "",
            "lastStampAt": _日時の文字(m.last_stamp_at) if m.last_stamp_at else "",
            "stampAchievedDate": (_日時の文字(m.stamp_achieved_at)
                                  if m.stamp_achieved_at else ""),
        }

    return {"status": "ok", "rewardStatus": 返す}


# ---------------------------------------------------------------------------
# お名前・フリガナの整え方
#
# **GAS の normalizeStoredName_ / normalizeStoredKana_ と同じにすること。**
# ここがずれると、同じ方が二重に登録される。
# ---------------------------------------------------------------------------

_空白 = re.compile(r"[\s\u3000]+")


def _名前を整える(値):
    """空白を全部取る。

    「山田 太郎」と「山田太郎」が別人として登録され、
    同じ方が二重に登録されてしまうため（GAS 6870行のコメントと同じ理由）。
    """
    return _空白.sub("", "" if 値 is None else str(値))


def _かなを整える(値):
    """NFKCで寄せ、空白を取り、**カタカナをひらがなにする。**

    ふりがなは「ひらがなで登録していただく」方針にしたが、
    以前に登録された分はカタカナのまま残っている。照合も保存もひらがなに揃える。
    """
    t = "" if 値 is None else str(値)
    t = unicodedata.normalize("NFKC", t)
    t = _空白.sub("", t)
    # ァ(30A1)〜ヶ(30F6) を 0x60 引いてひらがなへ
    return "".join(
        chr(ord(c) - 0x60) if "\u30a1" <= c <= "\u30f6" else c for c in t
    )


# GAS の normalizeRegistrationSource_ と同じ言い換え表。
_登録経路の言い換え = {
    "new": "新規登録", "newregistration": "新規登録", "register": "新規登録",
    "recover": "復元", "identity": "復元",
    "transfer": "引き継ぎコード利用", "transfercode": "引き継ぎコード利用",
    "merge": "重複候補からの復旧", "duplicate": "重複候補からの復旧",
    "activity": "自動復旧", "auto": "自動復旧",
}


def _登録経路(値):
    s = str(値 or "").strip()
    if not s:
        return "新規登録"
    return _登録経路の言い換え.get(s.lower(), s)


def _日付(値, 既定):
    """生年月日。**読めなければ既定値（いまの値）を残す。**空にしない。"""
    s = str(値 or "").strip()
    if not s:
        # はっきり空を送ってきたときだけ、空にする。
        return None if 値 is not None and str(値) == "" else 既定
    d = parse_date(s[:10])
    if d is not None:
        return d
    t = parse_datetime(s)
    return t.date() if t is not None else 既定


def _通知が入か(届け先):
    """届け先の中身から、通知オンかどうかを決める。

    アプリは `pushSubscription: false`（真偽値）か、購読IDの文字列を送ってくる。
    `false` は「通知オフ」。**文字列の "false" も同じ扱いにする。**
    """
    if 届け先 is False or 届け先 is None:
        return False
    s = str(届け先).strip()
    if not s or s.lower() in ("false", "0", "none", "null", "undefined"):
        return False
    return True


def 会員を書き換える(data):
    """GAS の handleUpdateUser()（6759行）と同じ。**新規登録も兼ねる。**

    お客様がプロフィールを保存したとき、通知の入切を変えたときに呼ばれる。
    **会員が無ければ作る。**行を作るのは、この口だけの仕事。

    ## 送られてこなかった項目は、いまの値を残す

    JSON では `undefined` の項目は**そもそも送られてこない**ので、
    「キーがあるか」で判断する。GAS の `data.x !== undefined` と同じ。

        キーが無い      → いまの値のまま
        キーがあって "" → **空にする**（はっきり消したいということ）

    ## GAS と違うところ

    1. **パスコードはハッシュにして保存する。**GAS は平文で表に書いている。
       平文をやめるのが、この移行の目的の1つ。照合は check_password で行う。
    2. **スタンプは、読めない値で 0 にしない。**理由は 特典をそろえる() と同じ。
    3. **行ロックを取る。**GAS は読んで書くだけで、同時に届くと片方が消える。

    ## 気をつけていること

    お名前とフリガナは**保存する前に整える**（空白を取る・ひらがなに寄せる）。
    ここがGASとずれると、同じ方が二重に登録される。
    """
    d = data or {}
    会員ID = str(d.get("memberId") or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "IDが指定されていません"}

    # GAS と同じ検査。お名前を送ってきたのに、整えると空になる場合は断る。
    if "name" in d and not _名前を整える(d.get("name")):
        return {"status": "error", "message": "お名前を入力してください"}

    いま = timezone.now()

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(member_id=会員ID).first()
        新規 = m is None
        if 新規:
            # **ここが唯一、会員を作る場所。**
            m = Member(member_id=会員ID, created_at=いま)
            m.device_sessions = []
            m.registration_source = _登録経路(d.get("registrationSource"))
            m.registration_source_detail = str(d.get("registrationSourceDetail") or "").strip()
            m.registration_source_updated_at = いま

        if "name" in d:
            m.name = _名前を整える(d.get("name"))
        else:
            # GAS は既存の値も normalizeStoredName_ に通し直している。
            # 空白入りで登録された古い記録が、保存のたびに整う。同じにする。
            m.name = _名前を整える(m.name)

        m.kana = _かなを整える(d["kana"] if "kana" in d else m.kana)

        if "phone" in d:
            m.phone = str(d.get("phone") or "").strip()
        if "address" in d:
            m.address = str(d.get("address") or "").strip()
        if "avatar" in d:
            m.avatar_url = str(d.get("avatar") or "").strip()
        if "status" in d:
            m.status = str(d.get("status") or "").strip()
        if "memo" in d:
            # **お客様アプリは、保存のたびに memo:'' を送ってくる。**
            # つまり受付の覚え書きが毎回消える。GAS も同じ動きをしている。
            # ここは GAS に合わせてある。直すならアプリ側（app.js）で
            # memo を送るのをやめるのが正しい。
            m.memo = str(d.get("memo") or "")
        if "birthday" in d:
            m.birthday = _日付(d.get("birthday"), m.birthday)

        if "pushSubscription" in d:
            届け先 = d.get("pushSubscription")
            m.push_subscription = "" if 届け先 in (False, None) else str(届け先).strip()
            if m.push_subscription.lower() == "false":
                m.push_subscription = ""
            m.push_enabled = _通知が入か(届け先)

        if "passcode" in d and str(d.get("passcode") or "").strip():
            # **平文では保存しない。**GAS との違い。
            m.passcode_hash = make_password(str(d.get("passcode")).strip())
            # GAS と同じ。パスコードを決め直したら、引き継ぎコードは無効にする。
            m.transfer_code = ""
            m.transfer_code_issued_at = None

        if ("registrationSource" in d or "registrationSourceDetail" in d
                or "registrationSourceUpdatedAt" in d):
            m.registration_source = _登録経路(
                d.get("registrationSource") or m.registration_source)
            if "registrationSourceDetail" in d:
                m.registration_source_detail = str(
                    d.get("registrationSourceDetail") or "").strip()
            t = parse_datetime(str(d.get("registrationSourceUpdatedAt") or ""))
            m.registration_source_updated_at = (
                t if t is not None else (m.registration_source_updated_at or いま))

        # スタンプ・特典。**読めない値で 0 にしない。**
        if "stampCount" in d:
            m.stamp_count = _数(d.get("stampCount"), m.stamp_count or 0, 0, 10)
        if "stampCardNum" in d:
            m.stamp_card_number = _数(d.get("stampCardNum"), m.stamp_card_number or 1, 1)
        if "rewards" in d:
            m.reward_history = _一覧(d.get("rewards"), m.reward_history or [])
        if "stampHistory" in d:
            m.stamp_history = _一覧(d.get("stampHistory"), m.stamp_history or [])

        # **端末・退会の印・統合先には触らない。**GAS も持ち越すだけ。
        m.save()

    return {"status": "ok", "created": 新規}


# ---------------------------------------------------------------------------
# ごほうびガチャ
# ---------------------------------------------------------------------------

# GAS の REWARD_GACHA_PRIZE_POOL（管理者・お客様.js）と**同じにすること。**
# 色や文言まで写してあるのは、アプリがこの値をそのまま画面に出すから。
_賞の基本 = [
    {"key": "A", "rankLabel": "A賞", "capsuleColor": "#f5cb6c", "accentColor": "#b0791b",
     "message": "受付でその時の特典をお受け取りください。", "weight": 5},
    {"key": "B", "rankLabel": "B賞", "capsuleColor": "#f3b7c9", "accentColor": "#b86282",
     "message": "まゆみ助産院からのうれしいごほうびです。受付でご案内します。", "weight": 15},
    {"key": "C", "rankLabel": "C賞", "capsuleColor": "#b9d8a7", "accentColor": "#628f58",
     "message": "やさしいプレゼントをご用意しています。受付へお声がけください。", "weight": 30},
    {"key": "D", "rankLabel": "D賞", "capsuleColor": "#b9d9f3", "accentColor": "#547fa2",
     "message": "お楽しみプレゼントをご用意しています。", "weight": 50},
]


def _その月の賞(月キー):
    """その月の賞の中身と確率を出す。

    GAS の getRewardGachaMonthlyConfig_ と同じ選び方をする。

        その月の設定があれば、それ
        無ければ、**その月より前でいちばん新しいもの**
        それも無ければ、最後の1つ
        設定そのものが無ければ、既定の確率（5/15/30/50）

    「無ければ既定に落ちる」ではなく「前の月を引き継ぐ」のが要点。
    月が変わった瞬間に賞の中身が消えないため。
    """
    from apps.records.models import AppSetting

    設定 = AppSetting.objects.filter(pk="REWARD_GACHA_CONFIG").first()
    生 = (設定.value if 設定 else None) or {}
    if isinstance(生, str):
        try:
            生 = json.loads(生)
        except ValueError:
            生 = {}
    月ごと = 生.get("monthlyPrizes") if isinstance(生, dict) else None
    月ごと = 月ごと if isinstance(月ごと, list) else []

    選んだ = None
    直前 = None
    for e in 月ごと:
        if not isinstance(e, dict):
            continue
        if e.get("month") == 月キー:
            選んだ = e
            break
        if str(e.get("month") or "") <= 月キー:
            直前 = e
    選んだ = 選んだ or 直前 or (月ごと[-1] if 月ごと else None)
    賞 = (選んだ or {}).get("prizes") or {}

    出 = []
    for 基本 in _賞の基本:
        保存 = 賞.get(基本["key"]) or {}
        中身 = str(保存.get("content") or "").strip()
        確率 = 保存.get("probability")
        try:
            重み = max(0.0, min(100.0, round(float(確率) * 10) / 10))
        except (TypeError, ValueError):
            重み = float(基本["weight"])
        出.append(dict(
            基本,
            # GAS の formatRewardGachaRewardName_ と同じ。
            rewardName=(基本["rankLabel"] + " " + 中身) if 中身 else (基本["rankLabel"] + "プレゼント"),
            note=str(保存.get("note") or "").strip(),
            weight=重み,
        ))
    return 出


def _賞を引く(賞たち):
    """重み付きで1つ選ぶ。GAS の pickWeightedRewardGachaPrize_ と同じ。

    **重みの合計が0なら、最後の1つ（D賞）を返す。**
    全部0のときに「当たらない」で止めると、お客様が回せなくなる。
    """
    合計 = sum(max(0.0, p["weight"]) for p in 賞たち)
    if 合計 <= 0:
        return 賞たち[-1]
    出た = random.random() * 合計
    for p in 賞たち:
        出た -= max(0.0, p["weight"])
        if 出た < 0:
            return p
    return 賞たち[-1]


def _賞の見た目(賞名):
    """特典の名前から、色や文言を引き当てる。GAS の getRewardGachaPrizeMeta_ と同じ。"""
    名 = str(賞名 or "").strip()
    for p in _賞の基本:
        if 名 == p["rankLabel"] or (名 and 名.startswith(p["rankLabel"])):
            return p
    return {
        "key": "SPECIAL", "rankLabel": "ごほうび獲得",
        "capsuleColor": "#d9c5a2", "accentColor": "#8d6c46",
        "message": "受付でその時の特典をお受け取りください。",
    }


def _一か月後(t):
    """1か月後。**JavaScript の `setMonth(+1)` と同じにする。**

    JS は、はみ出した日数を翌月へ送る。

        1/31 → 2/31 は無いので **3/3**（うるう年でなければ）
        3/31 → 4/31 は無いので **5/1**

    「月末に丸める」ほうが自然に見えるが、**そうするとGASと有効期限が
    1〜3日ずれる。**お客様に見える日付なので、GASに合わせる。
    """
    import calendar
    import datetime as _dt

    年 = t.year + (1 if t.month == 12 else 0)
    月 = 1 if t.month == 12 else t.month + 1
    その月の末日 = calendar.monthrange(年, 月)[1]
    if t.day <= その月の末日:
        return t.replace(year=年, month=月)
    はみ出し = t.day - その月の末日
    return t.replace(year=年, month=月, day=その月の末日) + _dt.timedelta(days=はみ出し)


def _特典の返し方(特典, すでに引いた):
    見た目 = _賞の見た目(特典.get("rewardName"))
    return {
        "key": 見た目["key"],
        "rankLabel": 見た目["rankLabel"],
        "rewardName": str(特典.get("rewardName") or "特典プレゼント"),
        # GAS も実質 `reward.rewardNote || ''`（賞の基本表に note は無い）。
        "rewardNote": str(特典.get("rewardNote") or ""),
        "earnedDate": 特典.get("earnedDate") or "",
        "expiryDate": 特典.get("expiryDate") or "",
        "capsuleColor": 見た目["capsuleColor"],
        "accentColor": 見た目["accentColor"],
        "message": 見た目["message"],
        "alreadyDrawn": bool(すでに引いた),
    }


def _特典の状態(m):
    return {
        "stampCount": m.stamp_count or 0,
        "stampCardNum": max(1, m.stamp_card_number or 1),
        "rewards": m.reward_history or [],
        "stampHistory": m.stamp_history or [],
        "lastStampDate": m.last_stamp_at.astimezone().strftime("%Y-%m-%d") if m.last_stamp_at else "",
        "lastStampAt": _日時の文字(m.last_stamp_at) if m.last_stamp_at else "",
        "stampAchievedDate": _日時の文字(m.stamp_achieved_at) if m.stamp_achieved_at else "",
    }


def ガチャを引く(data):
    """GAS の handleDrawRewardGacha()（7552行）と同じ。

    スタンプが10個たまった方が、1枚のカードにつき**1回だけ**回せる。

    ## ここは「2回引かれない」ことが全て

    GAS は `LockService.getScriptLock()` で**スクリプト全体**を止めている。
    こちらは**その会員の行だけ**を止める（`select_for_update`）。
    他のお客様の通信は止まらないので、GASより速くて、同じだけ安全。

    二重に引かれない仕組みは、ロックだけに頼らない。

        いまのカード番号と同じ番号の特典が、すでにあるか？
          → あれば **引かずに、その特典をそのまま返す**（alreadyDrawn）

    **これが本当の守り。**通信が2回届いても、2つ目は「もう引いてあります」
    になるだけで、特典は増えない。ロックは、その判定と書き込みの間に
    割り込まれないようにするためのもの。

    ## 有効期限

    「スタンプが10個そろった日時」＋1か月。**回した日ではない。**
    そろった日時が記録に無ければ、いまの時刻を使う（GASも同じ）。
    """
    d = data or {}
    会員ID = str(d.get("memberId") or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが指定されていません"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(member_id=会員ID).first()
        if not m:
            return {"status": "error", "message": "会員情報が見つかりません"}

        いまのカード = max(1, m.stamp_card_number or 1)
        特典たち = m.reward_history if isinstance(m.reward_history, list) else []

        # **すでに引いてあれば、引かない。**ここが二重引きの本当の守り。
        for r in 特典たち:
            if not isinstance(r, dict):
                continue
            try:
                番号 = max(1, int(r.get("cardNum") or 1))
            except (TypeError, ValueError):
                番号 = 1
            if 番号 == いまのカード:
                return {
                    "status": "ok",
                    "alreadyDrawn": True,
                    "rewardStatus": _特典の状態(m),
                    "drawnReward": _特典の返し方(r, True),
                }

        if (m.stamp_count or 0) < 10:
            # GAS と同じ文言。
            return {"status": "error", "message": "スタンプが10個たまっていません"}

        獲得 = m.stamp_achieved_at or timezone.now()
        月キー = timezone.now().astimezone().strftime("%Y-%m")
        当たり = _賞を引く(_その月の賞(月キー))

        新しい特典 = {
            # GAS と同じ形。時刻＋乱数で、まず重ならない。
            "id": "reward-%d-%d" % (int(獲得.timestamp() * 1000), random.randint(0, 999)),
            "cardNum": いまのカード,
            "rewardName": 当たり["rewardName"],
            "rewardNote": 当たり["note"],
            "earnedDate": _日時の文字(獲得),
            "expiryDate": _日時の文字(_一か月後(獲得)),
            "used": False,
        }

        m.reward_history = [新しい特典] + 特典たち
        if not m.stamp_achieved_at:
            m.stamp_achieved_at = 獲得
        m.save(update_fields=["reward_history", "stamp_achieved_at", "changed_at"])

        返す = {
            "status": "ok",
            "rewardStatus": _特典の状態(m),
            "drawnReward": _特典の返し方(新しい特典, False),
        }

    return 返す


# action の名前 → 書き込む関数
書けること = {
    "syncUserDeviceSession": 端末をそろえる,
    "removeUserDeviceSession": 端末を外す,
    "syncUserRewardStatus": 特典をそろえる,
    "updateUser": 会員を書き換える,
    "drawRewardGacha": ガチャを引く,
}


def 受け取る(request):
    """POST を受ける。GAS の doPost と同じ形。

    アプリは `postToGAS({type: '...', ...})` で JSON を送ってくる。
    """
    if request.method != "POST":
        return None, {"status": "error", "message": "POST で送ってください。"}

    try:
        data = json.loads(request.body.decode("utf-8") or "{}")
    except ValueError:
        return None, {"status": "error", "message": "中身を読み取れませんでした。"}

    if not isinstance(data, dict):
        return None, {"status": "error", "message": "中身の形が違います。"}

    種類 = str(data.get("type") or "").strip()
    書く = 書けること.get(種類)
    if not 書く:
        # **黙って成功にしない。**まだ無い口だと分かるように返す。
        return None, {
            "status": "error",
            "message": f"'{種類}' はまだサーバー側にありません。GASをお使いください。",
            "notImplemented": True,
        }
    return 書く, data
