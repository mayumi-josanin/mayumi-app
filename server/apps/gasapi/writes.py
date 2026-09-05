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

import base64
import binascii
import hashlib
import json
import random
import re
import secrets
import unicodedata
from pathlib import Path

from django.conf import settings

from django.contrib.auth.hashers import check_password, make_password
from django.core.cache import cache
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
            # **お名前が無いなら、新しい会員は作らない。**
            #
            # GAS は作る。`data.name` が**そもそも送られてこなければ**検査を
            # 通り抜け、お名前も電話番号も生年月日も空の行ができる
            # （管理者・お客様.js 6779行）。
            #
            # **この口は合鍵なしで誰でも呼べる。**会員IDを適当に付けて送れば、
            # いくらでも空の会員が作れる。2026-08-23 に、切り替えの確認中に
            # 自分で1件作ってしまって気づいた（すぐ消した）。
            #
            # 2026年8月に「入れないお客様が21名」いた原因も、おそらくこれ。
            # 空の行はお名前も生年月日も無いので、**復元では永久に通れない。**
            #
            # 既にいる会員の書き換えは、お名前なしでも通す。
            # 通知の入切だけを送ってくる経路（app.js 5341行）が
            # `{memberId, pushSubscription}` しか送らないため。
            if not _名前を整える(d.get("name")):
                return {"status": "error", "message": "お名前を入力してください"}

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
            # **保存する前に、消えた先頭の0を戻す。**
            m.phone = _電話を整える(d.get("phone"))
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


# ---------------------------------------------------------------------------
# 入り口（アプリに戻っていただくための2つ）
#
# **ここがいちばん危ない。**通れば、その会員の記録すべてが手に入る。
# 通しすぎれば他人が入り、締めすぎれば本当のご本人が入れなくなる。
#
# 2026年8月に「入れないお客様が21名」いた。生年月日が必須一致だったため。
# **条件は GAS と1つも変えない。**変えるなら、それは別の判断として相談する。
# ---------------------------------------------------------------------------

# GAS と同じ。10分のあいだに10回はずしたら、いったんお休みいただく。
復元の上限 = 10
復元のお休み秒 = 10 * 60


def _電話を寄せる(値):
    """**照合するため**に数字だけにする。9桁以上で0で始まらなければ0を足す（GASと同じ）。

    保存するときは `_電話を整える()` を使う。こちらは比べるためのもので、
    ハイフンや空白を落としてしまうため、そのまま保存すると見た目が変わる。
    """
    数 = re.sub(r"\D", "", str(値 or "").strip())
    if len(数) >= 9 and not 数.startswith("0"):
        数 = "0" + 数
    return 数


def _電話を整える(値):
    """**保存する**電話番号。消えた先頭の0を戻す。見た目は変えない。

    スプレッドシートが電話番号を数値として受け取ると先頭の0が落ちる。
    2026-08-25 に既存を直したが、**新しい行にはまた落ちていた**（2026-08-27）。
    GAS 側にも同じ `normalizePhoneForStore_` を入れてある。
    **片方だけに頼ると、また同じことが起きる。**

    足すのは、足した結果が日本の電話番号の形になるときだけ。
    形にならないものは**そのまま持つ**（推測で作り変えない）。
    """
    文 = str(値 or "").strip()
    if not 文:
        return ""
    数 = re.sub(r"\D", "", 文)
    if not 数 or 数.startswith("0"):
        return 文

    # 国番号81（例: 817055600662 → 07055600662）
    if len(数) == 12 and 数.startswith("81"):
        return "0" + 数[2:]

    候補 = "0" + 数
    # 携帯（070/080/090 の11桁）
    if len(候補) == 11 and 候補[:3] in ("070", "080", "090"):
        return 候補
    # 固定電話（0で始まる10桁。07x は固定の市外局番に無い）
    if len(候補) == 10 and 候補[1] != "0" and not 候補.startswith("07"):
        return 候補

    return 文


def _日付を寄せる(値):
    """`YYYY-MM-DD` の文字にする。GAS の normalizeDateOnlyValue_ と同じ。"""
    if 値 is None or 値 == "":
        return ""
    if hasattr(値, "strftime"):
        return 値.strftime("%Y-%m-%d")
    文 = str(値).strip()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", 文)
    if m:
        return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    d = parse_date(文[:10])
    return d.strftime("%Y-%m-%d") if d else ""


def _引き継ぎコードを寄せる(値):
    return re.sub(r"\D", "", str(値 or ""))[:8]


def _復元の鍵(名, かな):
    """試行回数を数えるための鍵。**お名前そのものは残さない。**

    GAS と同じく SHA-256 にして頭だけ使う。
    記録に平文のお名前が残ると、それ自体が漏れうる情報になる。
    """
    種 = str(名 or "") + "|" + str(かな or "")
    return "recover_" + hashlib.sha256(種.encode("utf-8")).hexdigest()[:24]


def _お休み中か(名, かな):
    try:
        return int(cache.get(_復元の鍵(名, かな)) or 0) >= 復元の上限
    except Exception:
        # **数えられないときに復元を止めない。**
        # 止めると、本当のご本人が締め出される。GASも同じ判断。
        return False


def _はずれを数える(名, かな):
    try:
        鍵 = _復元の鍵(名, かな)
        cache.set(鍵, int(cache.get(鍵) or 0) + 1, 復元のお休み秒)
    except Exception:
        pass


def _数えるのをやめる(名, かな):
    try:
        cache.delete(_復元の鍵(名, かな))
    except Exception:
        # 消せなくても10分で自然に消える。
        pass


def _会員の返し方(m, 平文パスコード=""):
    """GAS の buildRecoverAccountUserFromRow_ と同じ形。

    **`rewards` と `stampHistory` は「文字列」で返す。**
    GAS がセルの中身（JSONの文字列）をそのまま返しているため。
    配列で返すとアプリ側の読み取りが変わる。
    """
    return {
        "memberId": m.member_id,
        "name": m.name or "",
        "kana": m.kana or "",
        "phone": m.phone or "",
        "avatar": m.avatar_url or "",
        "memo": m.memo or "",
        "status": m.status or "",
        "birthday": _日付を寄せる(m.birthday),
        "address": m.address or "",
        "passcode": 平文パスコード,
        "deviceSessions": m.device_sessions or [],
        "stampCount": m.stamp_count or 0,
        "stampCardNum": max(1, m.stamp_card_number or 1),
        "rewards": json.dumps(m.reward_history or [], ensure_ascii=False),
        "stampHistory": json.dumps(m.stamp_history or [], ensure_ascii=False),
        "lastStampDate": m.last_stamp_at.astimezone().strftime("%Y-%m-%d") if m.last_stamp_at else "",
        "lastStampAt": _日時の文字(m.last_stamp_at) if m.last_stamp_at else "",
        "stampAchievedAt": _日時の文字(m.stamp_achieved_at) if m.stamp_achieved_at else "",
        "regDate": m.created_at.astimezone().strftime("%Y-%m-%d") if m.created_at else "",
    }


def _入り直していただく(m, 新パスコード, 経路, 詳細):
    """通ったあとの共通の後始末。

    **退会の印を消す。**復元は「戻ってきていただく」手続きなので、
    消えた扱いになっていた方も、ここで元に戻る（GASも同じ）。
    """
    m.passcode_hash = make_password(新パスコード)
    m.deleted = False
    m.deleted_at = None
    m.merged_into_id = ""
    m.transfer_code = ""
    m.transfer_code_issued_at = None
    m.registration_source = 経路
    m.registration_source_detail = 詳細
    m.registration_source_updated_at = timezone.now()
    m.save()


def 会員を復元する(data):
    """GAS の handleRecoverAccount()（7202行）と同じ。

    端末を替えた方・アプリを消してしまった方が、記録に戻るための入り口。

    ## 通る条件（**GASと1つも変えていない**）

    引き継ぎコードがあれば、それだけで通る（ご本人しか持たないため）。

    無ければ、

        生年月日が**完全に一致**すること（必須）
        かつ、お名前**または**フリガナのどちらかが一致すること
        電話番号・いまのパスコードは、**入力されたときだけ**絞り込みに使う

    1件も無ければ断る。**2件以上でも断る**（取り違えないため）。

    > 2026年8月に「入れないお客様が21名」いたのは、生年月日が必須だから。
    > **条件を緩めるかどうかは、この移行とは別の判断。**ここでは変えない。

    ## GAS と違うところ

    1. **パスコードはハッシュで照合する**（`check_password`）。GASは平文比較。
    2. **行ロックを取る。**
    3. 試行回数は、GASのキャッシュではなく**データベース**で数える。
       gunicorn が3つ動いているので、処理ごとに数えると10回が30回になる。
    """
    d = data or {}
    名 = _名前を整える(d.get("name"))
    かな = _かなを整える(d.get("kana"))
    電話 = _電話を寄せる(d.get("phone"))
    生年月日 = _日付を寄せる(d.get("birthday"))
    パスコード = str(d.get("passcode") or "").strip()
    新パスコード = str(d.get("newPasscode") or "").strip() or パスコード
    引き継ぎ = _引き継ぎコードを寄せる(d.get("transferCode"))

    # **GASと同じ順で確かめる。**順が違うと、出るお知らせが変わる。
    if not 引き継ぎ and not 名 and not かな:
        return {"status": "error", "message": "お名前またはフリガナを入力してください。"}
    if not 引き継ぎ and not 生年月日:
        return {"status": "error", "message": "生年月日を入力してください。"}
    if not re.match(r"^(?:\d{4}|\d{6})$", 新パスコード):
        return {"status": "error",
                "message": "この端末で使うパスコードを4桁または6桁の数字で入力してください。"}

    # 引き継ぎコードはご本人しか持たないので数えない。
    if not 引き継ぎ and _お休み中か(名, かな):
        return {"status": "error",
                "message": "復元の試行が続いたため、一時的に停止しています。10分後に再度お試しください。"}

    with transaction.atomic():
        if 引き継ぎ:
            m = (Member.objects.select_for_update()
                 .filter(transfer_code=引き継ぎ).first())
            if not m:
                return {"status": "error",
                        "message": "一致する会員情報が見つかりませんでした。入力内容をご確認ください。"}
            if not m.transfer_code_issued_at:
                # **使えないコードは、その場で消す。**GASも同じ。
                m.transfer_code = ""
                m.save(update_fields=["transfer_code", "changed_at"])
                return {"status": "error",
                        "message": "この引き継ぎコードは無効です。元の端末で新しいコードを発行してください。"}
            期限 = m.transfer_code_issued_at + timezone.timedelta(hours=168)  # 1週間
            if 期限 < timezone.now():
                m.transfer_code = ""
                m.transfer_code_issued_at = None
                m.save(update_fields=["transfer_code", "transfer_code_issued_at", "changed_at"])
                return {"status": "error",
                        "message": "この引き継ぎコードは期限切れです。元の端末で新しいコードを発行してください。"}
            _入り直していただく(m, 新パスコード, "引き継ぎコード利用", "引き継ぎコードで復元")
            return {"status": "ok",
                    "user": dict(_会員の返し方(m), passcode=新パスコード),
                    "recoveredBy": "transferCode"}

        # ここから、ご本人の情報で探す経路。
        # **退会の印がある方も探す。**復元は「戻ってきていただく」手続きなので。
        候補 = []
        for m in Member.objects.select_for_update():
            if _日付を寄せる(m.birthday) != 生年月日:
                continue
            名が一致 = bool(名) and _名前を整える(m.name) == 名
            かなが一致 = bool(かな) and bool(m.kana) and _かなを整える(m.kana) == かな
            if not 名が一致 and not かなが一致:
                continue
            if 電話 and _電話を寄せる(m.phone) != 電話:
                continue
            if パスコード:
                # **ハッシュで照合する。**GASは平文比較だった。
                if not m.passcode_hash or not check_password(パスコード, m.passcode_hash):
                    continue
            候補.append(m)

        if not 候補:
            _はずれを数える(名, かな)
            return {"status": "error",
                    "message": "一致する会員情報が見つかりませんでした。入力内容をご確認ください。"}
        if len(候補) > 1:
            # 複数一致は正しい会員でも起きるので、はずれとして数えない（GASと同じ）。
            return {"status": "error",
                    "message": "一致する会員情報が複数見つかりました。電話番号または現在のパスコードを追加で入力するか、引き継ぎコードをご利用ください。"}

        _数えるのをやめる(名, かな)
        m = 候補[0]
        _入り直していただく(m, 新パスコード, "復元", "本人情報で復元")
        return {"status": "ok",
                "user": dict(_会員の返し方(m), passcode=新パスコード),
                "recoveredBy": "identity"}


def パスコードを付け直す(data):
    """GAS の handleResetForgottenPasscode()（7377行）と同じ。

    パスコードを忘れた方が、付け直すための入り口。

    ## 通る条件（**GASと1つも変えていない**）

        お名前が一致すること（必須）
        かつ、**記録にある項目だけ**を照合する

    記録に電話番号があるなら、電話番号も一致しないと通らない。
    記録に生年月日があるなら、生年月日も一致しないと通らない。
    **記録が空の項目は、確かめようがないので問わない。**

    ただし「お名前だけで通す」ことはしない（同姓同名の他人が入れてしまう）。
    **記録に電話も生年月日も無い方は、この画面からは通れない。**受付でご対応。

    > 電話番号と生年月日の「どちらか一方でよい」のは、記録の側が空の方が
    > 4割近くいるため。両方必須にすると、その方々が原理的に通れなくなる。
    """
    d = data or {}
    会員ID = str(d.get("memberId") or "").strip()
    名 = _名前を整える(d.get("name"))
    電話 = _電話を寄せる(d.get("phone"))
    生年月日 = _日付を寄せる(d.get("birthday"))
    新パスコード = str(d.get("newPasscode") or "").strip()

    if not 名 or not 新パスコード or (not 電話 and not 生年月日):
        return {"status": "error",
                "message": "お名前と、電話番号または生年月日のどちらか、そして新しいパスコードを入力してください。"}
    if not re.match(r"^(?:\d{4}|\d{6})$", 新パスコード):
        return {"status": "error",
                "message": "新しいパスコードは4桁または6桁の数字で入力してください。"}
    if _お休み中か(名, ""):
        return {"status": "error",
                "message": "お手続きが続いたため、一時的にお休みしています。10分ほど経ってから、もう一度お試しください。"}

    def 見る(m):
        """通せるなら「いくつ照合できたか」を返す。通せないなら None。"""
        if _名前を整える(m.name) != 名:
            return None
        行の電話 = _電話を寄せる(m.phone)
        行の生年月日 = _日付を寄せる(m.birthday)
        数 = 0
        if 行の電話:
            if not 電話 or 行の電話 != 電話:
                return None
            数 += 1
        if 行の生年月日:
            if not 生年月日 or 行の生年月日 != 生年月日:
                return None
            数 += 1
        if 数 == 0:
            # 記録に電話も生年月日も無い方は、受付でご対応。
            return None
        return 数

    with transaction.atomic():
        候補 = []
        for m in Member.objects.select_for_update():
            数 = 見る(m)
            if 数:
                候補.append((m, 数))

        選んだ = None
        if 会員ID:
            # 会員IDの指定があれば、その方に絞る（取り違えを防ぐ）。
            選んだ = next((c for c in 候補 if c[0].member_id == 会員ID), None)
        if not 選んだ:
            if len(候補) > 1:
                _はずれを数える(名, "")
                return {"status": "error",
                        "message": "同じ内容のご登録が複数見つかりました。恐れ入りますが、受付にお申し出ください。"}
            選んだ = 候補[0] if 候補 else None

        if not 選んだ:
            _はずれを数える(名, "")
            return {"status": "error",
                    "message": "ご登録の内容と一致しませんでした。"
                               "ご登録時に電話番号や生年月日をいただいていない場合は、"
                               "この画面からはお手続きできません。"
                               "恐れ入りますが、受付にお申し出ください。"}

        _数えるのをやめる(名, "")
        m, 照合数 = 選んだ
        _入り直していただく(m, 新パスコード, m.registration_source or "新規登録",
                            m.registration_source_detail or "")

        # 何で通したかを残す。**パスコードそのものは書かない。**
        _パスコード再設定を記録する(m, 照合数)

    return {"status": "ok", "user": dict(_会員の返し方(m), passcode=新パスコード)}


def _パスコード再設定を記録する(m, 照合数):
    """GAS の recordPasscodeResetHistory_ と同じ。操作履歴に残す。"""
    try:
        from apps.records.models import AuditLog

        AuditLog.objects.create(
            # **シートの行番号は付けない。**サーバーが作った記録の印になる。
            sheet_row=None,
            happened_at=timezone.now(),
            kind="パスコード再設定",
            result="成功",
            target=m.member_id,
            summary=("氏名・電話番号・生年月日の3点で確認" if 照合数 >= 2
                     else "氏名と、記録にある1項目で確認（もう一方は未登録）"),
            operator="ご本人",
            detail=None,
        )
    except Exception:
        # **記録が残せなくても、お客様の手続きは止めない。**
        pass


# action の名前 → 書き込む関数
def 写真を受け取る(d):
    """GAS の handleUploadImage（管理者・お客様.js 5491行）と同じ形で答える。

        受け取る  {type:'uploadImage', filename, mimeType, base64}
        返す      {status:'ok', url:…, fileId:…}

    アプリは返ってきた `url` を、そのまま会員の「アイコンURL」に入れる。
    **だから URL の形が違っても構わないが、返さないと写真が消える。**

    ## GAS との違い

    GAS は Google Drive に上げ、**リンクを知れば誰でも見られる**共有URLを返す。
    ここでは同じ見え方に揃えてある（切り替えで見え方まで変えると、写真が
    出ないときに原因を切り分けられなくなるため）。
    ただし**ファイル名に乱数を入れて推測できないようにしてある。**
    会員IDから名前が決まると、他の方の写真を順に見て回れてしまう。

    ## 受け付ける大きさ

    アプリは写真を縮めてから送ってくるが、**信じない。**
    大きすぎるものは断る。ディスクを埋められると、予約も記録も止まる。
    """
    生 = d.get("base64") or ""
    if not 生:
        return {"status": "error", "message": "画像がありませんでした。"}

    # data URL の頭が付いたまま来ることがある。落としてから読む。
    文 = str(生)
    if "," in 文[:64] and 文[:5] == "data:":
        文 = 文.split(",", 1)[1]

    try:
        中身 = base64.b64decode(文, validate=True)
    except (binascii.Error, ValueError):
        return {"status": "error", "message": "画像を読み取れませんでした。"}

    if not 中身:
        return {"status": "error", "message": "画像が空でした。"}
    if len(中身) > 8 * 1024 * 1024:
        return {"status": "error", "message": "画像が大きすぎます。8MBまでにしてください。"}

    # **中身を見て種類を決める。**送られてきた mimeType は信じない。
    # 拡張子だけで判断すると、画像でないものを置かれる。
    先頭 = 中身[:12]
    if 先頭[:3] == b"\xff\xd8\xff":
        拡張子 = "jpg"
    elif 先頭[:8] == b"\x89PNG\r\n\x1a\n":
        拡張子 = "png"
    elif 先頭[:6] in (b"GIF87a", b"GIF89a"):
        拡張子 = "gif"
    elif 先頭[:4] == b"RIFF" and 先頭[8:12] == b"WEBP":
        拡張子 = "webp"
    else:
        return {"status": "error", "message": "画像として読めない形式でした。"}

    名前 = secrets.token_hex(16) + "." + 拡張子
    置き場 = Path(settings.MEDIA_ROOT) / "avatars"
    置き場.mkdir(parents=True, exist_ok=True)
    (置き場 / 名前).write_bytes(中身)

    return {
        "status": "ok",
        "url": (settings.PUBLIC_BASE_URL.rstrip("/") + "/media/avatars/" + 名前),
        # GAS は Drive のファイルIDを返している。アプリは使っていないが、
        # **形をそろえておく。**返ってこない項目があると、あとで困る。
        "fileId": 名前,
    }


def 通知の届け先を外す(d):
    """GAS の handleUnsubscribePush と同じ。通知を切ったときに呼ばれる。

    **届け先を空にするだけ。**会員の他の項目には触らない。

    2026-09-05 時点で、届け先を持つ会員は**0名**。動いていない経路だが、
    「通知を切りたい」に応えられないのは、機能の欠落ではなく信頼の問題。
    """
    会員ID = str(d.get("memberId") or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが必要です"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(pk=会員ID).first()
        if not m:
            return {"status": "error", "message": "会員が見つかりませんでした"}
        m.push_subscription = ""
        m.push_enabled = False
        m.save(update_fields=["push_subscription", "push_enabled"])

    return {"status": "ok", "message": "通知を止めました"}


def _お知らせの窓口(名):
    """お知らせの窓口を呼ぶ。**お知らせ以外の表なら、まだ無いと答える。**"""
    def 呼ぶ(d):
        from . import admin_news

        if not admin_news.お知らせ宛てか(d):
            return {"status": "error", "notImplemented": True,
                    "message": "お知らせ以外の表は、まだサーバーにありません。"}
        return getattr(admin_news, 名)(d)
    return 呼ぶ


_お知らせ足す = _お知らせの窓口("足す")
_お知らせ書き換える = _お知らせの窓口("書き換える")
_お知らせ公開 = _お知らせの窓口("公開を変える")
_お知らせ一覧掲載 = _お知らせの窓口("一覧掲載を変える")
_お知らせ一覧から外す = _お知らせの窓口("一覧から外す")
_お知らせ消す = _お知らせの窓口("消す")
_お知らせまとめて消す = _お知らせの窓口("まとめて消す")


def _FAQ保存(d):
    from . import admin_faq

    return admin_faq.保存(d)


def _FAQ消す(d):
    from . import admin_faq

    return admin_faq.消す(d)


def _カテゴリ(名):
    def 呼ぶ(d):
        from . import admin_category

        return getattr(admin_category, 名)(d)
    return 呼ぶ


_カテゴリ足す = _カテゴリ("足す")
_カテゴリ書き換える = _カテゴリ("書き換える")
_カテゴリ消す = _カテゴリ("消す")


書けること = {
    "syncUserDeviceSession": 端末をそろえる,
    "removeUserDeviceSession": 端末を外す,
    "syncUserRewardStatus": 特典をそろえる,
    "updateUser": 会員を書き換える,
    "drawRewardGacha": ガチャを引く,
    "recoverAccount": 会員を復元する,
    "resetForgottenPasscode": パスコードを付け直す,
    "uploadImage": 写真を受け取る,
    "unsubscribePush": 通知の届け先を外す,
    # 管理アプリ向け（お知らせの表）。**公開アクションに入れない**ので合鍵が要る。
    # deleteRow / updateRecordStatus は表をまたぐ窓口なので、
    # お知らせ以外が来たら notImplemented を返す（GAS がシートで処理する）。
    "addBlog": _お知らせ足す,
    "updateBlog": _お知らせ書き換える,
    "updateRecordStatus": _お知らせ公開,
    "updateNoticeVisibility": _お知らせ一覧掲載,
    "deleteNoticeListing": _お知らせ一覧から外す,
    "deleteRow": _お知らせ消す,
    "deleteRows": _お知らせまとめて消す,
    # 使い方FAQ。表をまたがない窓口なので、振り分けの守りは要らない。
    "saveSupportFaq": _FAQ保存,
    "deleteSupportFaq": _FAQ消す,
    # カテゴリ。表をまたがない。
    "addCategory": _カテゴリ足す,
    "updateCategory": _カテゴリ書き換える,
    "deleteCategory": _カテゴリ消す,
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
