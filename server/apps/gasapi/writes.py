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

from django.db import transaction
from django.utils import timezone

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
            "stampAchievedDate": "",
        }

    return {"status": "ok", "rewardStatus": 返す}


# action の名前 → 書き込む関数
書けること = {
    "syncUserDeviceSession": 端末をそろえる,
    "removeUserDeviceSession": 端末を外す,
    "syncUserRewardStatus": 特典をそろえる,
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
