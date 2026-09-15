"""管理側の窓口（合鍵が要る）。受付が使う。

    getAdminUsers             管理者・お客様.js 8482行
    getPushUsers              同 8795行
    checkMemberToken          同 2260行（会員を見に行く部分が checkMemberOnServer）
    updateAdminUser           同 8272行
    deleteUser                同 8375行
    updateAdminRewardStatus   同 8395行
    grantSurveyStamp          同 4571行
    issueTransferCode         同 7823行
    mergeUsers                同 8652行
    setUserPasscode / stopUserPush   GAS には無い。移行後にシートを直せない代わり
"""

import datetime as dt
from zoneinfo import ZoneInfo

import pytest
from django.contrib.auth.hashers import check_password
from django.utils import timezone
from django.utils.dateparse import parse_datetime

from apps.gasapi import admin_member
from apps.members.models import Member
from apps.records.models import OrderLine
from tests.gasapi_member_support import 会員を作る, 書く, 読み直す, 読む

pytestmark = [pytest.mark.django_db, pytest.mark.usefixtures("会員試験の下ごしらえ")]

JST = ZoneInfo("Asia/Tokyo")


def 日本時間(*a):
    return dt.datetime(*a, tzinfo=JST)


def 管理で書く(client, 中身):
    return 書く(client, 中身, 合鍵あり=True).json()


def 管理で読む(client, action, 引数=None):
    return 読む(client, action, 引数, 合鍵あり=True).json()


# ═════════════════════════════════════════════════════════
# getAdminUsers
# ═════════════════════════════════════════════════════════

# GAS 8511行が返す項目。rowIdx はシートの行番号なので、サーバーには無い
# （管理アプリは memberId を先に見て探す。admin/index.html 7339行）。
一覧の項目 = {
    "memberId", "timestamp", "name", "kana", "phone", "avatarUrl", "memo", "pushEnabled",
    "status", "birthday", "address", "deviceSessions", "deviceCount",
    "stampCount", "stampCardNum", "rewards", "stampHistory",
    "lastStampDate", "lastStampAt", "stampAchievedDate", "latestStampActivityAt",
    "registrationSource", "registrationSourceDetail", "registrationSourceUpdatedAt",
    "orderCount", "pendingOrderCount", "lastOrderAt", "orderTotal",
    "surveyAnsweredAt", "surveyStampGrantedAt", "surveyStampPendingAt",
}


def test_会員一覧_項目がそろっている(client):
    会員を作る(name="佐藤花子", kana="さとうはなこ", phone="09012345678", memo="覚え書き",
            birthday="1990-04-05", address="平塚市", status="産後", stamp_count=3,
            stamp_card_number=1, push_subscription="sub-abc",
            device_sessions=[{"deviceId": "d1"}], created_at=日本時間(2026, 1, 15, 10, 0),
            last_stamp_at=日本時間(2026, 4, 1, 10, 0),
            survey_answered_at=日本時間(2026, 4, 2, 9, 0),
            survey_stamp_granted_at=日本時間(2026, 4, 2, 9, 1))
    答 = 管理で読む(client, "getAdminUsers")
    assert 答["status"] == "ok" and len(答["users"]) == 1
    u = 答["users"][0]
    assert set(u) == 一覧の項目
    assert u["memberId"] == "MYM-1001" and u["name"] == "佐藤花子" and u["kana"] == "さとうはなこ"
    assert u["phone"] == "09012345678" and u["memo"] == "覚え書き" and u["status"] == "産後"
    assert u["address"] == "平塚市"
    assert u["deviceCount"] == 1 and u["deviceSessions"] == [{"deviceId": "d1"}]
    assert u["stampCount"] == 3 and u["stampCardNum"] == 1
    assert u["lastStampDate"] == "2026-04-01"
    assert parse_datetime(u["lastStampAt"]) == 日本時間(2026, 4, 1, 10, 0)
    assert parse_datetime(u["latestStampActivityAt"]) == 日本時間(2026, 4, 1, 10, 0)
    assert parse_datetime(u["timestamp"]) == 日本時間(2026, 1, 15, 10, 0)
    assert parse_datetime(u["surveyAnsweredAt"]) == 日本時間(2026, 4, 2, 9, 0)
    assert parse_datetime(u["surveyStampGrantedAt"]) == 日本時間(2026, 4, 2, 9, 1)
    assert u["surveyStampPendingAt"] == ""
    assert u["orderCount"] == 0 and u["orderTotal"] == 0 and u["lastOrderAt"] == ""


def test_会員一覧_通知は届け先の有無で真偽値(client):
    """一覧の pushEnabled は真偽値でよい（GAS 8523行 !!row[PUSH]）。届け先そのものは getPushUsers。"""
    会員を作る("MYM-1001", push_subscription="sub-abc")
    会員を作る("MYM-1002", push_subscription="")
    一覧 = {u["memberId"]: u["pushEnabled"] for u in 管理で読む(client, "getAdminUsers")["users"]}
    assert 一覧 == {"MYM-1001": True, "MYM-1002": False}


def test_会員一覧_退会の印がある方は出ない(client):
    会員を作る("MYM-1001")
    会員を作る("MYM-1002", deleted=True)
    assert [u["memberId"] for u in 管理で読む(client, "getAdminUsers")["users"]] == ["MYM-1001"]


def test_会員一覧_お客様に見えないメモが管理には出る(client):
    会員を作る(memo="お体のこと")
    assert 管理で読む(client, "getAdminUsers")["users"][0]["memo"] == "お体のこと"


def test_会員一覧_注文の集計は注文の表から作る(client):
    会員を作る("MYM-1001")
    OrderLine.objects.create(order_id="O-1", member_id="MYM-1001", product_name="お茶",
                             subtotal=500, status="受付中", ordered_at=timezone.now())
    u = 管理で読む(client, "getAdminUsers")["users"][0]
    assert u["orderCount"] == 1 and u["pendingOrderCount"] == 1 and u["orderTotal"] == 500
    assert u["lastOrderAt"] != ""


def test_会員一覧_カード番号が0の方は1で返す(client):
    会員を作る(stamp_card_number=0)
    assert 管理で読む(client, "getAdminUsers")["users"][0]["stampCardNum"] == 1


# ═════════════════════════════════════════════════════════
# getPushUsers — 届け先そのもの（GAS 8795行）
# ═════════════════════════════════════════════════════════

届け先の項目 = {"memberId", "name", "phone", "birthday", "status", "subscription", "stampCount",
             "rewardCount", "orderCount", "pendingOrderCount", "lastOrderAt", "deviceCount"}


def test_届け先一覧_届け先そのものを返し真偽値にしない(client):
    会員を作る("MYM-1001", name="佐藤花子", phone="09012345678", birthday="1990-04-05", status="産後",
            push_subscription="abc123-subscription-id", stamp_count=4,
            reward_history=[{"id": "r1", "used": False}, {"id": "r2", "used": True}, {"id": "r3"}],
            device_sessions=[{"deviceId": "d1"}, {"deviceId": "d2"}])
    会員を作る("MYM-1002", push_subscription="true")
    会員を作る("MYM-1003", push_subscription="")
    会員を作る("MYM-1004", push_subscription="sub-x", deleted=True)
    答 = 管理で読む(client, "getPushUsers")
    assert 答["status"] == "ok"
    assert [u["memberId"] for u in 答["users"]] == ["MYM-1001", "MYM-1002"]
    u = 答["users"][0]
    assert set(u) == 届け先の項目
    assert u["subscription"] == "abc123-subscription-id"
    assert isinstance(u["subscription"], str) and u["subscription"] not in ("true", "false")
    assert u["name"] == "佐藤花子" and u["phone"] == "09012345678"
    assert u["birthday"] == "1990-04-05" and u["status"] == "産後"
    assert u["stampCount"] == 4
    assert u["rewardCount"] == 2          # 使っていない特典の数
    assert u["deviceCount"] == 2
    assert 答["users"][1]["subscription"] == "true"


def test_届け先一覧_誰もいなければ空(client):
    assert 管理で読む(client, "getPushUsers") == {"status": "ok", "users": []}


# ═════════════════════════════════════════════════════════
# checkMemberOnServer — checkMemberToken の会員を見に行く部分（GAS 2269行）
# ═════════════════════════════════════════════════════════

def test_札の会員確認_いる会員は名前とフリガナだけ返す(client):
    会員を作る(name="佐藤花子", kana="さとうはなこ", phone="09012345678")
    答 = 管理で読む(client, "checkMemberOnServer", {"memberId": "MYM-1001"})
    assert 答 == {"status": "ok", "valid": True, "memberId": "MYM-1001",
                  "name": "佐藤花子", "kana": "さとうはなこ"}


def test_札の会員確認_退会者はvalidがfalse(client):
    会員を作る(deleted=True)
    assert 管理で読む(client, "checkMemberOnServer", {"memberId": "MYM-1001"}) == {"status": "ok", "valid": False}


def test_札の会員確認_いない会員や空のIDもvalidがfalse(client):
    assert 管理で読む(client, "checkMemberOnServer", {"memberId": "MYM-9999"}) == {"status": "ok", "valid": False}
    assert 管理で読む(client, "checkMemberOnServer", {}) == {"status": "ok", "valid": False}


# ═════════════════════════════════════════════════════════
# updateAdminUser（GAS 8272行）
# ═════════════════════════════════════════════════════════

def test_会員更新_送った項目だけ変わる(client):
    m = 会員を作る(name="佐藤花子", memo="覚え書き", phone="09012345678", stamp_count=3,
                push_subscription="sub-abc", transfer_code="12345678")
    答 = 管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1001", "rowIdx": 5,
                          "name": "佐藤華子", "birthday": "1990/04/05", "status": "産後",
                          "role": "管理者", "bijirisRegistered": True, "stampCount": "5",
                          "lineUserId": "U-abc"})
    assert 答 == {"status": "ok"}
    m = 読み直す(m)
    assert m.name == "佐藤華子" and str(m.birthday) == "1990-04-05"
    assert m.status == "産後" and m.role == "管理者" and m.bijiris_registered is True
    assert m.stamp_count == 5 and m.line_user_id == "U-abc"
    # 送らなかった項目は残る
    assert m.memo == "覚え書き" and m.phone == "09012345678"
    # 受け付けない項目には触らない
    assert m.push_subscription == "sub-abc" and m.transfer_code == "12345678"


def test_会員更新_お名前を空にできない(client):
    m = 会員を作る(name="佐藤花子")
    答 = 管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1001", "name": " "})
    assert 答 == {"status": "error", "message": "お名前は空にできません"}
    assert 読み直す(m).name == "佐藤花子"


def test_会員更新_電話の先頭の0を戻し空の生年月日は空にする(client):
    m = 会員を作る(birthday="1990-04-05")
    管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1001", "phone": "9012345678", "birthday": ""})
    m = 読み直す(m)
    assert m.phone == "09012345678" and m.birthday is None


def test_会員更新_LINEのIDは2人に付けない(client):
    会員を作る("MYM-1001", line_user_id="U-abc")
    m = 会員を作る("MYM-1002")
    答 = 管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1002", "lineUserId": "U-abc"})
    assert 答 == {"status": "error", "message": "そのLINEユーザーIDは、ほかの会員（MYM-1001）に付いています。"}
    assert 読み直す(m).line_user_id is None
    # 空にして外すのは許す
    管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1001", "lineUserId": ""})
    assert Member.objects.get(pk="MYM-1001").line_user_id is None


def test_会員更新_いない会員(client):
    答 = 管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-9999", "name": "x"})
    assert 答 == {"status": "error", "message": "会員が見つかりません"}


def test_会員更新_お名前の空白を取りフリガナをひらがなに寄せて保存する(client):
    m = 会員を作る()
    管理で書く(client, {"type": "updateAdminUser", "memberId": "MYM-1001", "name": "佐藤 花子", "kana": "サトウ　ハナコ"})
    m = 読み直す(m)
    assert m.name == "佐藤花子" and m.kana == "さとうはなこ"


# ═════════════════════════════════════════════════════════
# deleteUser — 印だけ。行は残る（GAS 8375行）
# ═════════════════════════════════════════════════════════

def test_会員削除_印を付けるだけで行は残る(client):
    m = 会員を作る(name="佐藤花子", memo="覚え書き")
    答 = 管理で書く(client, {"type": "deleteUser", "memberId": "MYM-1001", "rowIdx": 5})
    assert 答 == {"status": "ok"}
    m = 読み直す(m)
    assert m.deleted is True and m.deleted_at is not None
    assert m.name == "佐藤花子" and m.memo == "覚え書き"
    assert Member.objects.count() == 1
    # 一覧・届け先・札の確認から消える
    assert 管理で読む(client, "getAdminUsers")["users"] == []
    assert 管理で読む(client, "checkMemberOnServer", {"memberId": "MYM-1001"})["valid"] is False


def test_会員削除_いない会員(client):
    答 = 管理で書く(client, {"type": "deleteUser", "memberId": "MYM-9999"})
    assert 答 == {"status": "error", "message": "会員が見つかりません"}


# ═════════════════════════════════════════════════════════
# updateAdminRewardStatus — 受付がスタンプを直す（GAS 8395行）
# ═════════════════════════════════════════════════════════

def test_特典の書き換え_値が変わり管理者が直した時刻が残る(client):
    m = 会員を作る(stamp_count=2, last_stamp_at=timezone.now())
    答 = 管理で書く(client, {"type": "updateAdminRewardStatus", "rowIdx": 5, "memberId": "MYM-1001",
                          "stampCount": 8, "stampCardNum": 2, "lastStampDate": "2026-04-01",
                          "stampAchievedDate": "2026-04-01T10:00:00+09:00",
                          "rewards": [{"id": "r1", "cardNum": 1, "used": True}]})
    assert 答["status"] == "ok"
    m = 読み直す(m)
    assert m.stamp_count == 8 and m.stamp_card_number == 2
    assert m.reward_history == [{"id": "r1", "cardNum": 1, "used": True}]
    assert timezone.localtime(m.stamp_achieved_at) == 日本時間(2026, 4, 1, 10, 0)
    assert m.reward_admin_set_at is not None
    # お客様アプリはこれを見て「サーバーを正とする」（views.py 特典の状態を見る）
    状態 = 読む(client, "getUserRewardStatus", {"memberId": "MYM-1001"}).json()
    assert 状態["adminSetAt"] != "" and 状態["rewardStatus"]["stampCount"] == 8


def test_特典の書き換え_デモ用の取得制限の解除(client):
    m = 会員を作る(stamp_count=2, last_stamp_at=timezone.now())
    管理で書く(client, {"type": "updateAdminRewardStatus", "memberId": "MYM-1001",
                     "stampCount": 10, "clearLastStampDate": True})
    m = 読み直す(m)
    assert m.stamp_count == 10 and m.last_stamp_at is None


def test_特典の書き換え_いない会員(client):
    答 = 管理で書く(client, {"type": "updateAdminRewardStatus", "memberId": "MYM-9999", "stampCount": 1})
    assert 答 == {"status": "error", "message": "会員が見つかりません"}


def test_特典の書き換え_返す形はGASと同じ(client):
    会員を作る(stamp_count=2)
    答 = 管理で書く(client, {"type": "updateAdminRewardStatus", "memberId": "MYM-1001", "stampCount": 3})
    assert 答["rewardStatus"]["stampCount"] == 3
    assert set(答["rewardStatus"]) == {"stampCount", "stampCardNum", "rewards", "stampHistory",
                                       "lastStampDate", "lastStampAt", "stampAchievedDate"}


# ═════════════════════════════════════════════════════════
# grantSurveyStamp — 会員ごとの印で二重に付けない（GAS 4571行）
# ═════════════════════════════════════════════════════════

def お礼(client, 会員ID="MYM-1001"):
    return 管理で書く(client, {"type": "grantSurveyStamp", "memberId": 会員ID})


def test_お礼スタンプ_1個付いて印が残る(client):
    m = 会員を作る(stamp_count=3)
    答 = お礼(client)
    assert 答 == {"status": "ok", "answered": True, "granted": True, "stampCount": 4}
    m = 読み直す(m)
    assert m.stamp_count == 4 and m.last_stamp_at is not None
    assert m.survey_answered_at is not None and m.survey_stamp_granted_at is not None
    assert m.survey_stamp_pending_at is None
    u = 管理で読む(client, "getAdminUsers")["users"][0]
    assert u["surveyAnsweredAt"] != "" and u["surveyStampGrantedAt"] != ""


def test_お礼スタンプ_二重に付かない(client):
    """印（SURVEY_STAMP:<会員ID>）を移し忘れると二重に付く（CLAUDE.md）。"""
    m = 会員を作る(stamp_count=3)
    お礼(client)
    答 = お礼(client)
    assert 答 == {"status": "ok", "answered": True, "granted": False, "reason": "already_granted"}
    assert 読み直す(m).stamp_count == 4


def test_お礼スタンプ_取り込んだ印があれば付けない(client):
    """GAS で付けた印（プロパティ）を取り込んだ方。サーバーで再び付けてはいけない。"""
    m = 会員を作る(stamp_count=3, survey_stamp_granted_at=日本時間(2026, 5, 1, 10, 0))
    答 = お礼(client)
    assert 答["granted"] is False and 答["reason"] == "already_granted"
    assert 読み直す(m).stamp_count == 3


def test_お礼スタンプ_満杯なら保留にする(client):
    m = 会員を作る(stamp_count=10)
    答 = お礼(client)
    assert 答 == {"status": "ok", "answered": True, "granted": False, "reason": "card_full_pending"}
    m = 読み直す(m)
    assert m.stamp_count == 10
    assert m.survey_answered_at is not None            # 回答した事実は残す
    assert m.survey_stamp_pending_at is not None and m.survey_stamp_granted_at is None


def test_お礼スタンプ_いない会員と会員IDなし(client):
    assert お礼(client, "MYM-9999") == {"status": "error", "message": "会員が見つかりません: MYM-9999"}
    assert 管理で書く(client, {"type": "grantSurveyStamp"}) == {"status": "error", "message": "会員IDが指定されていません"}


def test_お礼スタンプ_履歴にも残る(client):
    m = 会員を作る(stamp_count=3, stamp_history=[{"acquiredDate": "2026-01-01T00:00:00+09:00"}])
    お礼(client)
    履歴 = 読み直す(m).stamp_history
    assert len(履歴) == 2 and 履歴[0]["note"] == "アンケート回答のお礼"


# ═════════════════════════════════════════════════════════
# issueTransferCode（GAS 7823行）
# ═════════════════════════════════════════════════════════

def test_引き継ぎコード_8桁で1週間有効(client):
    m = 会員を作る()
    前 = timezone.now()
    答 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"})
    assert 答["status"] == "ok"
    assert len(答["transferCode"]) == 8 and 答["transferCode"].isdigit()
    発行 = parse_datetime(答["issuedAt"])
    期限 = parse_datetime(答["expiresAt"])
    assert 期限 - 発行 == dt.timedelta(hours=168)
    assert 答["expiresAtLabel"]                                   # 入口が見せる文字
    m = 読み直す(m)
    assert m.transfer_code == 答["transferCode"]
    assert m.transfer_code_issued_at is not None and m.transfer_code_issued_at >= 前.replace(microsecond=0)


def test_引き継ぎコード_他の方と重ならない(client, monkeypatch):
    会員を作る("MYM-1001", transfer_code="00001234")
    会員を作る("MYM-1002")
    出す = iter([1234, 1234, 5678])
    monkeypatch.setattr(admin_member.random, "randrange", lambda *a, **k: next(出す))
    答 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1002"})
    assert 答["transferCode"] == "00005678"


def test_引き継ぎコード_発行し直すと前のは使えない(client):
    会員を作る()
    一 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"})["transferCode"]
    二 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"})["transferCode"]
    assert Member.objects.get(pk="MYM-1001").transfer_code == 二
    if 一 != 二:
        答 = 書く(client, {"type": "recoverAccount", "transferCode": 一, "newPasscode": "5678"}).json()
        assert 答["status"] == "error"


def test_引き継ぎコード_合鍵なしでも発行できる_お客様アプリも呼ぶ(client):
    """PUBLIC_ACTIONS に入っている（start/index.html 904行が会員IDで呼ぶ）。"""
    会員を作る()
    答 = 書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"}).json()
    assert 答["status"] == "ok"


def test_引き継ぎコード_発行してそのコードで復元できる(client):
    会員を作る(name="佐藤花子")
    符号 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"})["transferCode"]
    答 = 書く(client, {"type": "recoverAccount", "transferCode": 符号, "newPasscode": "5678"}).json()
    assert 答["status"] == "ok" and 答["user"]["memberId"] == "MYM-1001"


def test_引き継ぎコード_いない会員と会員IDなし(client):
    assert 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-9999"}) == {
        "status": "error", "message": "会員情報が見つかりませんでした。"}
    assert 管理で書く(client, {"type": "issueTransferCode"}) == {
        "status": "error", "message": "会員IDが指定されていません。"}


def test_引き継ぎコード_返す項目はGASと同じ(client):
    会員を作る()
    答 = 管理で書く(client, {"type": "issueTransferCode", "memberId": "MYM-1001"})
    assert set(答) == {"status", "transferCode", "issuedAt", "issuedAtLabel",
                       "expiresAt", "expiresAtLabel", "ttlHours"}
    assert 答["ttlHours"] == 168


# ═════════════════════════════════════════════════════════
# mergeUsers（GAS 8652行）— 他の表も動くこと
# ═════════════════════════════════════════════════════════

def 統合(client, 先="MYM-1001", 元=("MYM-1002",)):
    return 管理で書く(client, {"type": "mergeUsers", "targetMemberId": 先, "sourceMemberIds": list(元)})


def test_統合_空いている項目だけ埋め_スタンプは多いほう_特典と端末は合わせる(client):
    先 = 会員を作る("MYM-1001", name="佐藤花子", kana="", phone="", memo="先のメモ", stamp_count=3,
                 stamp_card_number=1, reward_history=[{"id": "r1"}], stamp_history=[{"a": 1}],
                 device_sessions=[{"deviceId": "d1"}], passcode="1111", deleted=True)
    元 = 会員を作る("MYM-1002", name="佐藤はなこ", kana="さとうはなこ", phone="09012345678",
                 memo="元のメモ", stamp_count=7, stamp_card_number=2,
                 reward_history=[{"id": "r1"}, {"id": "r2"}], stamp_history=[{"a": 2}],
                 device_sessions=[{"deviceId": "d1"}, {"deviceId": "d2"}], passcode="2222",
                 last_stamp_at=timezone.now())
    答 = 統合(client)
    assert 答["status"] == "ok"
    先 = 読み直す(先)
    assert 先.name == "佐藤花子" and 先.memo == "先のメモ"          # 上書きしない
    assert 先.kana == "さとうはなこ" and 先.phone == "09012345678"   # 空いていたところは埋まる
    assert check_password("1111", 先.passcode_hash)                # パスコードも上書きしない
    assert 先.stamp_count == 7 and 先.stamp_card_number == 2
    assert 先.reward_history == [{"id": "r1"}, {"id": "r2"}]       # 同じものは1つだけ
    assert 先.stamp_history == [{"a": 1}, {"a": 2}]
    assert [d["deviceId"] for d in 先.device_sessions] == ["d1", "d2"]
    assert 先.last_stamp_at is not None
    # 統合先の削除の印は外れ、経路が残る
    assert 先.deleted is False and 先.merged_into_id == ""
    assert 先.registration_source == "重複候補からの復旧"
    assert 先.registration_source_detail == "会員統合で情報を集約"
    # 元の会員は印だけ。行は残る
    元 = 読み直す(元)
    assert 元.deleted is True and 元.deleted_at is not None and 元.merged_into_id == "MYM-1001"
    assert 元.name == "佐藤はなこ"
    assert Member.objects.count() == 2
    # 一覧には統合先だけ
    assert [u["memberId"] for u in 管理で読む(client, "getAdminUsers")["users"]] == ["MYM-1001"]


def test_統合_すでに他へ統合済みの行は触らない(client):
    会員を作る("MYM-1001", phone="")
    元 = 会員を作る("MYM-1002", phone="09012345678", deleted=True, merged_into_id="MYM-3000")
    答 = 統合(client)
    assert 答["status"] == "ok"
    assert Member.objects.get(pk="MYM-1001").phone == ""
    assert 読み直す(元).merged_into_id == "MYM-3000"


def test_統合_自分自身は元にならない_いない元は飛ばす(client):
    会員を作る("MYM-1001")
    元 = 会員を作る("MYM-1002")
    答 = 統合(client, 元=("MYM-1001", "MYM-9999", "MYM-1002"))
    assert 答["status"] == "ok"
    assert Member.objects.get(pk="MYM-1001").deleted is False
    assert 読み直す(元).deleted is True


def test_統合_引数不足と統合先なし(client):
    assert 統合(client, 元=()) == {"status": "error", "message": "統合対象が不足しています"}
    assert 統合(client, 先="", 元=("MYM-1002",)) == {"status": "error", "message": "統合対象が不足しています"}
    会員を作る("MYM-1002")
    assert 統合(client, 先="MYM-9999") == {"status": "error", "message": "統合先会員が見つかりません"}


def test_統合_注文の会員IDも統合先に付け替える(client):
    会員を作る("MYM-1001")
    会員を作る("MYM-1002")
    OrderLine.objects.create(order_id="O-1", member_id="MYM-1002", product_name="お茶", subtotal=500)
    統合(client)
    assert OrderLine.objects.get(order_id="O-1").member_id == "MYM-1001"
    assert 管理で読む(client, "getAdminUsers")["users"][0]["orderCount"] == 1


# ═════════════════════════════════════════════════════════
# setUserPasscode / stopUserPush — GAS には無い。管理アプリが直接呼ぶ
# ═════════════════════════════════════════════════════════

def test_パスコード設定_新しいパスコードで入れるようになり旧方式は消える(client):
    m = 会員を作る(name="佐藤花子", passcode="1234", password_hash="old", password_salt="salt")
    答 = 管理で書く(client, {"type": "setUserPasscode", "memberId": "MYM-1001", "passcode": "5678"})
    assert 答 == {"status": "ok", "message": "パスコードを設定しました"}
    m = 読み直す(m)
    assert check_password("5678", m.passcode_hash) and not check_password("1234", m.passcode_hash)
    assert m.password_hash == "" and m.password_salt == ""
    assert 書く(client, {"type": "loginAccount", "name": "佐藤花子", "passcode": "5678"}).json()["status"] == "ok"


def test_パスコード設定_形が違えば断る(client):
    会員を作る()
    答 = 管理で書く(client, {"type": "setUserPasscode", "memberId": "MYM-1001", "passcode": "12345"})
    assert 答 == {"status": "error", "message": "パスコードは数字4桁または6桁で入力してください。"}


def test_通知を止める_届け先が空になる(client):
    m = 会員を作る(push_subscription="sub-abc", push_enabled=True)
    答 = 管理で書く(client, {"type": "stopUserPush", "memberId": "MYM-1001"})
    assert 答 == {"status": "ok", "message": "通知を止めました"}
    m = 読み直す(m)
    assert m.push_subscription == "" and m.push_enabled is False
