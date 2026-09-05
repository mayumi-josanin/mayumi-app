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


# ═════════════════════════════════════════════════════════
# 第2段：管理側の書き込み
#
# 受付が使う。**間違えても、その場で気づける。**
# `handleMergeUsers`（会員の統合）は他の表も触るので、まだ作っていない。
# ═════════════════════════════════════════════════════════

import random
import re

from django.db import transaction
from django.utils import timezone

引き継ぎコードの桁 = 8
引き継ぎコードの有効時間 = 168  # 1週間。GAS の TRANSFER_CODE_TTL_HOURS と同じ


def _電話を整える(文):
    """**消えた先頭の0を戻す。**GAS の normalizePhoneForStore_ と同じ考え方。

    シートは数字として扱うので `080…` が `80…` になる。
    2026-08 に121件を直した。**新しく入るぶんも同じように直す。**
    """
    値 = _文(文).strip()
    if not 値:
        return ""
    数 = re.sub(r"\D", "", 値)
    if not 数 or 数.startswith("0"):
        return 値
    if len(数) == 12 and 数.startswith("81"):
        return "0" + 数[2:]
    候補 = "0" + 数
    if len(候補) == 11 and 候補[:3] in ("070", "080", "090"):
        return 候補
    if len(候補) == 10 and 候補[1] != "0" and not 候補.startswith("07"):
        return 候補
    return 値


def _探す(d):
    """**会員IDで探す。**お名前では探さない（同姓同名・改名・表記ゆれ）。

    管理アプリは `rowIdx`（シートの行）でも指してくる。会員IDを優先し、
    無ければ行番号で探す。行番号はシートの並びなので、
    **サーバーでは会員IDのほうが確か。**
    """
    会員ID = _文(d.get("memberId")).strip()
    if 会員ID:
        return Member.objects.filter(pk=会員ID).first()
    return None


def 会員を書き換える(d):
    """GAS の handleUpdateAdminUser。**送られてこなかった項目は触らない。**"""
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "会員が見つかりません"}

    if "name" in d:
        名 = _文(d.get("name")).strip()
        if not 名:
            # **お名前を空にしない。**空の会員が作られた事故がある（2026-08-24）
            return {"status": "error", "message": "お名前は空にできません"}
        m.name = 名
    if "kana" in d:
        m.kana = _文(d.get("kana")).strip()
    if "phone" in d:
        m.phone = _電話を整える(d.get("phone"))
    if "memo" in d:
        m.memo = _文(d.get("memo"))
    if "birthday" in d:
        from django.utils.dateparse import parse_date

        文 = _文(d.get("birthday")).strip()
        m.birthday = parse_date(文[:10].replace("/", "-")) if 文 else None
    if "address" in d:
        m.address = _文(d.get("address")).strip()
    m.save()
    return {"status": "ok"}


def 特典を書き換える(d):
    """GAS の handleUpdateAdminRewardStatus。受付がスタンプを直すときに使う。"""
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "会員が見つかりません"}

    if "stampCount" in d:
        m.stamp_count = _数(d.get("stampCount"))
    if "stampCardNum" in d:
        m.stamp_card_number = _数(d.get("stampCardNum"))
    if "rewards" in d and isinstance(d.get("rewards"), list):
        m.reward_history = d["rewards"]
    if "stampHistory" in d and isinstance(d.get("stampHistory"), list):
        m.stamp_history = d["stampHistory"]
    if "lastStampAt" in d:
        m.last_stamp_at = _日時(d.get("lastStampAt"))
    if "stampAchievedDate" in d or "stampAchievedAt" in d:
        m.stamp_achieved_at = _日時(d.get("stampAchievedAt") or d.get("stampAchievedDate"))
    # **受付が直したときは、その日時を残す。**
    # お客様の端末が持っている記録と食い違ったとき、どちらが新しいかの判断に使う。
    m.reward_admin_set_at = timezone.now()

    if d.get("clearLastStampDate") is True:
        m.last_stamp_at = None
    m.save()
    return {"status": "ok", "stampCount": m.stamp_count or 0}


def _数(v, 既定=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 既定


def _日時(v):
    from django.utils.dateparse import parse_datetime

    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_datetime(文)
    if d and timezone.is_naive(d):
        d = timezone.make_aware(d)
    return d


def 会員を消す(d):
    """GAS の handleDeleteUser。**印を付けるだけ。行は残る。**

    消してしまうと、その方が復元しようとしたときに何も手がかりが無くなる。
    """
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "会員が見つかりません"}
    m.deleted = True
    m.deleted_at = timezone.now()
    m.save(update_fields=["deleted", "deleted_at", "changed_at"])
    return {"status": "ok"}


def お礼スタンプを付ける(d):
    """GAS の handleGrantSurveyStamp。アンケートに答えた方へのお礼。

    **二重に付けない。**すでに付けた印があれば何もしない。
    GAS はこの印をスクリプトプロパティに置いているが、
    サーバーでは会員の列（`survey_stamp_granted_at`）に持つ。
    """
    会員ID = _文(d.get("memberId")).strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが指定されていません"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(pk=会員ID).first()
        if not m:
            return {"status": "error", "message": f"会員が見つかりません: {会員ID}"}
        if m.survey_stamp_granted_at:
            return {"status": "ok", "answered": True, "granted": False,
                    "reason": "already_granted"}

        いま = timezone.now()
        if not m.survey_answered_at:
            m.survey_answered_at = いま

        # カードがいっぱいなら保留にする（GAS と同じ）
        if (m.stamp_count or 0) >= 10:
            m.survey_stamp_pending_at = いま
            m.save()
            return {"status": "ok", "answered": True, "granted": False,
                    "reason": "card_full_pending"}

        m.stamp_count = (m.stamp_count or 0) + 1
        m.last_stamp_at = いま
        m.survey_stamp_granted_at = いま
        m.survey_stamp_pending_at = None
        m.save()
    return {"status": "ok", "answered": True, "granted": True, "stampCount": m.stamp_count}


def 引き継ぎコードを出す(d):
    """GAS の handleIssueTransferCode。機種変更のときに使う。

    **8桁の数字。1週間有効。他の方と重ならないもの。**
    """
    会員ID = _文(d.get("memberId")).strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが指定されていません。"}

    with transaction.atomic():
        m = Member.objects.select_for_update().filter(pk=会員ID).first()
        if not m:
            return {"status": "error", "message": "会員情報が見つかりませんでした。"}

        使用中 = {
            re.sub(r"\D", "", _文(c))[:引き継ぎコードの桁]
            for c in Member.objects.exclude(transfer_code="").values_list("transfer_code", flat=True)
        }
        符号 = ""
        for _ in range(30):
            候補 = str(random.randrange(10 ** 引き継ぎコードの桁)).zfill(引き継ぎコードの桁)
            if 候補 not in 使用中:
                符号 = 候補
                break
        if not 符号:
            # GAS と同じ逃げ道。いまの時刻の下8桁。
            符号 = str(int(timezone.now().timestamp() * 1000))[-引き継ぎコードの桁:]

        発行 = timezone.now()
        m.transfer_code = 符号
        m.transfer_code_issued_at = 発行
        m.save(update_fields=["transfer_code", "transfer_code_issued_at", "changed_at"])

    from datetime import timedelta

    期限 = 発行 + timedelta(hours=引き継ぎコードの有効時間)
    return {
        "status": "ok",
        "transferCode": 符号,
        "issuedAt": _時刻の字(発行),
        "expiresAt": _時刻の字(期限),
        "expiresAtLabel": timezone.localtime(期限).strftime("%Y年%-m月%-d日 %H:%M"),
    }


def 会員を統合する(d):
    """GAS の handleMergeUsers。同じ方が2つの会員IDを持ってしまったときに使う。

    ## 何をするか（GAS 8367行と同じ）

    1. **統合先の空いている項目だけを、元の会員から埋める。**上書きはしない
    2. スタンプは**多いほう**を採る。特典と履歴と端末は**合わせる**
    3. 元の会員に**削除の印と「統合先」**を付ける（行は残す）
    4. 統合先の削除の印を**外す**（消えていた方へ統合することがある）
    5. 注文の会員IDを付け替える

    ## 5は、まだできない

    **注文の表をまだ移していない。**注文は0件なので今は害が無いが、
    注文を移すときに**ここも直す。**忘れると、統合した方の注文が
    元の会員IDのまま取り残される。

    ## お名前では探さない

    会員IDで探す。**同姓同名の方を統合したら、取り返しがつかない。**
    """
    統合先ID = _文(d.get("targetMemberId")).strip()
    元ID = d.get("sourceMemberIds")
    元ID = [_文(x).strip() for x in 元ID if _文(x).strip()] if isinstance(元ID, list) else []
    元ID = [x for x in 元ID if x != 統合先ID]
    if not 統合先ID or not 元ID:
        return {"status": "error", "message": "統合対象が不足しています"}

    埋める項目 = ["name", "kana", "phone", "avatar_url", "memo",
                  "push_subscription", "status", "birthday", "address", "passcode_hash"]

    with transaction.atomic():
        先 = Member.objects.select_for_update().filter(pk=統合先ID).first()
        if not 先:
            return {"status": "error", "message": "統合先会員が見つかりません"}

        統合した = []
        for mid in 元ID:
            m = Member.objects.select_for_update().filter(pk=mid).first()
            if not m:
                continue
            # すでに他へ統合済みの行は触らない（GAS と同じ）
            if m.deleted and _文(m.merged_into_id).strip():
                continue

            # ① 空いている項目だけ埋める。**上書きしない**
            for f in 埋める項目:
                いま = getattr(先, f, None)
                あちら = getattr(m, f, None)
                if not _文(いま).strip() and _文(あちら).strip():
                    setattr(先, f, あちら)

            # ② スタンプは多いほう
            先.stamp_count = max(先.stamp_count or 0, m.stamp_count or 0)
            先.stamp_card_number = max(先.stamp_card_number or 1, m.stamp_card_number or 1)
            先.reward_history = _合わせる(先.reward_history, m.reward_history)
            先.stamp_history = _合わせる(先.stamp_history, m.stamp_history)
            先.device_sessions = _合わせる(先.device_sessions, m.device_sessions)
            先.last_stamp_at = 先.last_stamp_at or m.last_stamp_at
            先.stamp_achieved_at = 先.stamp_achieved_at or m.stamp_achieved_at

            # ③ 元の会員に印を付ける。**行は残す**
            m.deleted = True
            m.deleted_at = timezone.now()
            m.merged_into_id = 統合先ID
            m.save()
            統合した.append(mid)

        # ④ 統合先の削除の印を外す
        先.deleted = False
        先.deleted_at = None
        先.merged_into_id = ""
        先.registration_source = "重複候補からの復旧"
        先.registration_source_detail = "会員統合で情報を集約"
        先.registration_source_updated_at = timezone.now()
        先.save()

    # ⑤ 注文の付け替えは、注文の表を移してから
    return {"status": "ok", "targetMemberId": 統合先ID, "merged": 統合した,
            "mergedCount": len(統合した)}


def _合わせる(a, b):
    """2つの並びを合わせる。**同じものは1つだけ。**順番は元のまま。

    中身が dict のこともあるので、比べるときは中身を並べ直した文字で見る。
    """
    import json as _json

    出, 見た = [], set()
    for 並び in (a or [], b or []):
        if not isinstance(並び, list):
            continue
        for x in 並び:
            try:
                鍵 = _json.dumps(x, ensure_ascii=False, sort_keys=True)
            except (TypeError, ValueError):
                鍵 = str(x)
            if 鍵 in 見た:
                continue
            見た.add(鍵)
            出.append(x)
    return 出
