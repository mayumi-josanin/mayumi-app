"""会員を、管理アプリから読むための窓口。**第1段（読むだけ）。**

計画は [会員の表を移す](docs/design/会員の表を移す.md)、
段取りは [表を移す手順](docs/design/表を移す手順.md)。

## 第1段は読むだけ

    getAdminUsers     管理アプリの会員一覧
    getPushUsers      **通知の届け先**
    checkMemberToken  ビジリスのGASが札を確かめに来る

**間違えても表示が変わるだけ。**書き込みは第2段・第3段で、
第1段が数日おかしくないことを見てから着手する。

## getPushUsers はとくに気をつける

`subscription` は**届け先そのもの。**真偽値にすると通知が届かなくなる。
2026-08-23 に取り込みの取りこぼしとして見つかっている
（購読IDを持つ方が False に落ちていた）。

## 注文の集計は 0 を返す

`orderCount` などは注文の表から作るが、**注文はまだ移していない**（0件）。
GAS 側も 0 を返すので、いまは同じ。**注文を移すときにここも直す。**
"""

from apps.members.models import Member

削除の印 = "削除済み"


def _文(v):
    return "" if v is None else str(v)


def _時刻の字(d):
    from django.utils import timezone

    if not d:
        return ""
    return timezone.localtime(d).strftime("%Y-%m-%dT%H:%M:%S+09:00")


def _日付の字(d):
    return d.strftime("%Y-%m-%d") if d else ""


def _特典の状態(m):
    """GAS の getRewardStatusFromRow_ と同じ形。"""
    return {
        "stampCount": m.stamp_count or 0,
        "stampCardNum": m.stamp_card_number or 0,
        "rewards": list(m.reward_history or []),
        "stampHistory": list(m.stamp_history or []),
        "lastStampDate": _日付の字(m.last_stamp_at.date() if m.last_stamp_at else None),
        "lastStampAt": _時刻の字(m.last_stamp_at),
        "stampAchievedDate": _日付の字(m.stamp_achieved_at.date() if m.stamp_achieved_at else None),
    }


def _注文の集計():
    """**注文はまだ移していない。**GAS も 0件なので同じ形の 0 を返す。"""
    return {"orderCount": 0, "pendingOrderCount": 0, "lastOrderAt": "", "orderTotal": 0}


def 一覧():
    """GAS の getAdminUsers。**消した会員は返さない。**"""
    出 = []
    for m in Member.objects.exclude(deleted=True).order_by("member_id"):
        特典 = _特典の状態(m)
        端末 = list(m.device_sessions or [])
        注文 = _注文の集計()
        # GAS は「スタンプの動きの、いちばん新しい日時」を出している
        最新 = max([x for x in (_時刻の字(m.last_stamp_at), _時刻の字(m.stamp_achieved_at)) if x] or [""])
        出.append({
            "memberId": m.member_id,
            "timestamp": _時刻の字(m.created_at),
            "name": m.name or "",
            "kana": m.kana or "",
            "phone": m.phone or "",
            "avatarUrl": m.avatar_url or "",
            "memo": m.memo or "",
            # **届け先が入っていれば「入」。**真偽値の列ではない。
            "pushEnabled": bool(_文(m.push_subscription).strip()) or bool(m.push_enabled),
            "status": m.status or "",
            "birthday": _日付の字(m.birthday),
            "address": m.address or "",
            "deviceSessions": 端末,
            "deviceCount": len(端末),
            "stampCount": 特典["stampCount"],
            "stampCardNum": 特典["stampCardNum"],
            "rewards": 特典["rewards"],
            "stampHistory": 特典["stampHistory"],
            "lastStampDate": 特典["lastStampDate"],
            "lastStampAt": 特典["lastStampAt"],
            "stampAchievedDate": 特典["stampAchievedDate"],
            "latestStampActivityAt": 最新,
            "registrationSource": m.registration_source or "",
            "registrationSourceDetail": m.registration_source_detail or "",
            "registrationSourceUpdatedAt": _時刻の字(m.registration_source_updated_at),
            "orderCount": 注文["orderCount"],
            "pendingOrderCount": 注文["pendingOrderCount"],
            "lastOrderAt": 注文["lastOrderAt"],
            "orderTotal": 注文["orderTotal"],
            "surveyAnsweredAt": _時刻の字(m.survey_answered_at),
            "surveyStampGrantedAt": _時刻の字(m.survey_stamp_granted_at),
            "surveyStampPendingAt": _時刻の字(m.survey_stamp_pending_at),
        })
    return {"status": "ok", "users": 出}


def 通知の届け先():
    """GAS の getPushUsers。**届け先を持つ方だけ。**

    `subscription` は届け先そのもの。**真偽値にしない。**
    ここを取り違えると、通知が1件も届かなくなる。
    """
    出 = []
    for m in Member.objects.exclude(deleted=True).order_by("member_id"):
        届け先 = _文(m.push_subscription).strip()
        if not 届け先:
            continue
        特典 = _特典の状態(m)
        注文 = _注文の集計()
        出.append({
            "memberId": m.member_id,
            "name": m.name or "",
            "phone": m.phone or "",
            "birthday": _日付の字(m.birthday),
            "status": m.status or "",
            "subscription": 届け先,
            "stampCount": 特典["stampCount"],
            # 使っていない特典の数
            "rewardCount": sum(1 for r in 特典["rewards"]
                               if isinstance(r, dict) and not r.get("used")),
            "orderCount": 注文["orderCount"],
            "pendingOrderCount": 注文["pendingOrderCount"],
            "lastOrderAt": 注文["lastOrderAt"],
            "deviceCount": len(list(m.device_sessions or [])),
        })
    return {"status": "ok", "users": 出}


def 札を確かめる(会員ID):
    """GAS の checkMemberToken のうち、会員を見に行く部分。

    **合っているかだけを返す。**個人を特定できる情報は名前とフリガナだけ
    （GAS がそう返しているため。増やさない）。
    """
    会員ID = _文(会員ID).strip()
    if not 会員ID:
        return {"status": "ok", "valid": False}
    m = Member.objects.filter(pk=会員ID).first()
    if not m or m.deleted:
        return {"status": "ok", "valid": False}
    return {"status": "ok", "valid": True, "memberId": m.member_id,
            "name": m.name or "", "kana": m.kana or ""}
