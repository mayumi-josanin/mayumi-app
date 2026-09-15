"""お客様側の窓口（合鍵なしで呼べるもの）。

    updateUser                管理者・お客様.js 7220行
    syncUserDeviceSession     同 8154行     getUserDevices           同 8175行
    removeUserDeviceSession   同 8194行
    syncUserRewardStatus      同 7995行     getUserRewardStatus      同 8218行
    drawRewardGacha           同 8039行
    recoverAccount            同 7685行     getRecoveryCandidates    同 7623行
    resetForgottenPasscode    同 7862行
    unsubscribePush           同 8106行（会員の行を空にする部分だけ）

**返す項目名・文言・空のときの挙動を GAS と同じにする。**アプリは返ってきた
中身で画面を描くので、1項目違うだけで画面が崩れる。
"""

import datetime as dt
import json
from zoneinfo import ZoneInfo

import pytest
from django.contrib.auth.hashers import check_password
from django.utils import timezone
from django.utils.dateparse import parse_datetime

from apps.gasapi import views
from apps.members.models import Member
from apps.records.models import AppSetting, AuditLog
from tests.gasapi_member_support import 会員を作る, 書く, 読み直す, 読む

pytestmark = [pytest.mark.django_db, pytest.mark.usefixtures("会員試験の下ごしらえ")]

JST = ZoneInfo("Asia/Tokyo")


def 日本時間(*a):
    return dt.datetime(*a, tzinfo=JST)


# ═════════════════════════════════════════════════════════
# updateUser — 新規登録も兼ねる（GAS 7220行）
# ═════════════════════════════════════════════════════════

def test_updateUser_お名前なしでは新しい会員を作れない(client):
    """空の会員が作られる穴（2026-08-23 に自分で1件作った）。GAS 7251行と同じ守り。"""
    答 = 書く(client, {"type": "updateUser", "memberId": "MYM-9999", "pushSubscription": "abc"}).json()
    assert 答 == {"status": "error", "message": "お名前を入力してください"}
    assert Member.objects.count() == 0


def test_updateUser_会員IDが無ければ断る(client):
    答 = 書く(client, {"type": "updateUser", "name": "佐藤花子"}).json()
    assert 答 == {"status": "error", "message": "IDが指定されていません"}


def test_updateUser_お名前があれば新しい会員を作る(client):
    答 = 書く(client, {"type": "updateUser", "memberId": "MYM-9999", "name": "佐藤 花子",
                     "kana": "サトウ　ハナコ", "phone": "9012345678", "birthday": "1990-04-05",
                     "address": "平塚市", "passcode": "1234", "stampCount": 2}).json()
    assert 答["status"] == "ok"
    m = Member.objects.get(pk="MYM-9999")
    # 空白を取り、ひらがなに寄せて保存（normalizeStoredName_ / normalizeStoredKana_）
    assert m.name == "佐藤花子"
    assert m.kana == "さとうはなこ"
    assert m.phone == "09012345678"           # 消えた先頭の0を戻す
    assert str(m.birthday) == "1990-04-05"
    assert m.registration_source == "新規登録"
    assert m.stamp_count == 2
    assert check_password("1234", m.passcode_hash)
    assert m.device_sessions == []


def test_updateUser_既存の会員は名前なしでも書き換えられる_他の項目を消さない(client):
    """通知の入切だけを送る経路（app.js 5352行）が {memberId, pushSubscription} しか送らない。"""
    m = 会員を作る(name="佐藤花子", kana="さとうはなこ", phone="09012345678",
                birthday="1990-04-05", address="平塚市", memo="受付の覚え書き",
                status="産後", stamp_count=4, avatar_url="https://x/a.png")
    答 = 書く(client, {"type": "updateUser", "memberId": "MYM-1001",
                     "pushSubscription": "sub-abc123"}).json()
    assert 答["status"] == "ok"
    m = 読み直す(m)
    assert m.push_subscription == "sub-abc123"
    assert m.push_enabled is True
    for 項目, 値 in [("name", "佐藤花子"), ("kana", "さとうはなこ"), ("phone", "09012345678"),
                   ("address", "平塚市"), ("memo", "受付の覚え書き"), ("status", "産後"),
                   ("stamp_count", 4), ("avatar_url", "https://x/a.png")]:
        assert getattr(m, 項目) == 値, 項目
    assert str(m.birthday) == "1990-04-05"


def test_updateUser_プロフィール保存はmemoを送らないので受付の覚え書きが残る(client):
    """app.js 7566行の形。memo は送らない（2026-09-05 に削除）。"""
    m = 会員を作る(memo="受付の覚え書き")
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "avatar": "", "name": "佐藤花子",
                 "phone": "09012345678", "birthday": "1990-04-05", "address": "平塚市",
                 "kana": "さとうはなこ", "pushSubscription": False})
    assert 読み直す(m).memo == "受付の覚え書き"


def test_updateUser_通知を切るとpushSubscriptionがfalseで届け先が空になる(client):
    m = 会員を作る(push_subscription="sub-abc", push_enabled=True)
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "pushSubscription": False})
    m = 読み直す(m)
    assert m.push_subscription == ""
    assert m.push_enabled is False


def test_updateUser_文字列のfalseも通知オフ(client):
    m = 会員を作る(push_subscription="sub-abc", push_enabled=True)
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "pushSubscription": "false"})
    m = 読み直す(m)
    assert m.push_subscription == ""
    assert m.push_enabled is False


def test_updateUser_お名前を空で送れば断る(client):
    m = 会員を作る(name="佐藤花子")
    答 = 書く(client, {"type": "updateUser", "memberId": "MYM-1001", "name": "　"}).json()
    assert 答 == {"status": "error", "message": "お名前を入力してください"}
    assert 読み直す(m).name == "佐藤花子"


def test_updateUser_パスコードを決め直すと引き継ぎコードは無効になる(client):
    """GAS 7305行 clearTransferCodeFromRow_。"""
    m = 会員を作る(transfer_code="12345678", transfer_code_issued_at=timezone.now())
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "passcode": "5678"})
    m = 読み直す(m)
    assert check_password("5678", m.passcode_hash)
    assert m.transfer_code == "" and m.transfer_code_issued_at is None


def test_updateUser_退会の印と端末には触らない(client):
    m = 会員を作る(deleted=True, merged_into_id="MYM-2000",
                device_sessions=[{"deviceId": "d1", "lastSeenAt": "2026-01-01T00:00:00+09:00"}])
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "name": "佐藤花子"})
    m = 読み直す(m)
    assert m.deleted is True and m.merged_into_id == "MYM-2000"
    assert m.device_sessions[0]["deviceId"] == "d1"


@pytest.mark.parametrize("値, 期待", [(None, 7), ("", 7), ("abc", 7), (9, 9), (0, 0), (99, 10)])
def test_updateUser_スタンプは読めない値で0に戻さない(client, 値, 期待):
    """**GAS と違えてある。**GAS は Number(x)||0 で null や空文字が 0 になり、
    お客様のスタンプが0に戻る（段階Bの下ごしらえ.md「読めない値が来ても、0にしない」）。
    「0にする」は管理アプリから明示的に行う。
    """
    m = 会員を作る(stamp_count=7)
    書く(client, {"type": "updateUser", "memberId": "MYM-1001", "stampCount": 値})
    assert 読み直す(m).stamp_count == 期待


# ═════════════════════════════════════════════════════════
# 端末 — syncUserDeviceSession / getUserDevices / removeUserDeviceSession
# ═════════════════════════════════════════════════════════

def 端末(i, 日=None):
    return {"deviceId": f"d-{i}", "label": f"端末{i}", "platform": "ios", "appVersion": "1.0",
            "lastSeenAt": f"2026-01-{日 or i:02d}T00:00:00+09:00",
            "passcodeEnabled": True, "pushEnabled": False, "current": False}


def test_端末の上限は8でGASと同じ():
    assert views.端末の上限 == 8


def test_端末同期_記録され返る(client):
    会員を作る()
    答 = 書く(client, {"type": "syncUserDeviceSession", "memberId": "MYM-1001", "deviceId": "d-1",
                     "label": "iPhone", "platform": "ios", "appVersion": "1.2.0",
                     "passcodeEnabled": True, "pushEnabled": "true"}).json()
    assert 答["status"] == "ok"
    assert len(答["devices"]) == 1
    d = 答["devices"][0]
    assert d["deviceId"] == "d-1" and d["label"] == "iPhone" and d["platform"] == "ios"
    assert d["appVersion"] == "1.2.0"
    assert d["current"] is True
    assert d["passcodeEnabled"] is True
    # === true のときだけ真。文字列の "true" は真にしない（GAS 4386行）
    assert d["pushEnabled"] is False
    assert parse_datetime(d["lastSeenAt"]) is not None
    assert set(d) == {"deviceId", "label", "platform", "appVersion", "lastSeenAt",
                      "passcodeEnabled", "pushEnabled", "current"}


def test_端末同期_同じ端末を2回送っても増えない(client):
    会員を作る()
    for _ in range(2):
        答 = 書く(client, {"type": "syncUserDeviceSession", "memberId": "MYM-1001",
                         "deviceId": "d-1"}).json()
    assert len(答["devices"]) == 1
    assert len(Member.objects.get(pk="MYM-1001").device_sessions) == 1


def test_端末同期_いま使っている端末だけcurrentが真(client):
    会員を作る(device_sessions=[dict(端末(1), current=True)])
    答 = 書く(client, {"type": "syncUserDeviceSession", "memberId": "MYM-1001", "deviceId": "d-2"}).json()
    assert [(d["deviceId"], d["current"]) for d in 答["devices"]] == [("d-2", True), ("d-1", False)]


def test_端末同期_9台目で最後に使ったのが古いものが落ちる(client):
    会員を作る(device_sessions=[端末(i) for i in range(1, 9)])
    答 = 書く(client, {"type": "syncUserDeviceSession", "memberId": "MYM-1001", "deviceId": "d-new"}).json()
    並び = [d["deviceId"] for d in 答["devices"]]
    assert len(並び) == 8
    assert 並び[0] == "d-new"
    assert "d-1" not in 並び               # いちばん古い d-1 が落ちる
    assert 並び[1:] == ["d-8", "d-7", "d-6", "d-5", "d-4", "d-3", "d-2"]
    assert len(Member.objects.get(pk="MYM-1001").device_sessions) == 8


@pytest.mark.parametrize("中身", [{"deviceId": "d-1"}, {"memberId": "MYM-1001"}])
def test_端末同期_会員IDか端末IDが無ければ断る(client, 中身):
    会員を作る()
    答 = 書く(client, dict(中身, type="syncUserDeviceSession")).json()
    assert 答 == {"status": "error", "message": "会員IDまたは端末IDが不足しています"}


def test_端末同期_いない会員には行を作らない(client):
    答 = 書く(client, {"type": "syncUserDeviceSession", "memberId": "MYM-9999", "deviceId": "d-1"}).json()
    assert 答 == {"status": "error", "message": "会員情報が見つかりません"}
    assert Member.objects.count() == 0


def test_端末一覧_いない会員はエラーではなく空(client):
    答 = 読む(client, "getUserDevices", {"memberId": "MYM-9999"}).json()
    assert 答 == {"status": "ok", "devices": []}


def test_端末一覧_会員IDが無ければエラー(client):
    答 = 読む(client, "getUserDevices", {}).json()
    assert 答 == {"status": "error", "message": "会員IDが必要です"}


def test_端末一覧_新しい順に8件まで_deviceIdの無いものは捨てる(client):
    記録 = [端末(i) for i in range(1, 11)] + [{"label": "壊れた記録"}, "文字列"]
    会員を作る(device_sessions=記録)
    答 = 読む(client, "getUserDevices", {"memberId": "MYM-1001"}).json()
    assert [d["deviceId"] for d in 答["devices"]] == [f"d-{i}" for i in range(10, 2, -1)]


def test_端末一覧_文字列のfalseを真にしない(client):
    会員を作る(device_sessions=[dict(端末(1), passcodeEnabled="false", pushEnabled="true", current="true")])
    d = 読む(client, "getUserDevices", {"memberId": "MYM-1001"}).json()["devices"][0]
    assert d["passcodeEnabled"] is False and d["pushEnabled"] is False and d["current"] is False


def test_端末を外す_その端末だけ消える_2回でも安全(client):
    会員を作る(device_sessions=[端末(1), 端末(2)])
    for _ in range(2):
        答 = 書く(client, {"type": "removeUserDeviceSession", "memberId": "MYM-1001", "deviceId": "d-1"}).json()
    assert 答["status"] == "ok"
    assert [d["deviceId"] for d in 答["devices"]] == ["d-2"]
    assert [d["deviceId"] for d in Member.objects.get(pk="MYM-1001").device_sessions] == ["d-2"]


def test_端末を外す_いない会員(client):
    答 = 書く(client, {"type": "removeUserDeviceSession", "memberId": "MYM-9999", "deviceId": "d-1"}).json()
    assert 答 == {"status": "error", "message": "会員情報が見つかりません"}


# ═════════════════════════════════════════════════════════
# 特典 — syncUserRewardStatus / getUserRewardStatus
# ═════════════════════════════════════════════════════════

特典の項目 = {"stampCount", "stampCardNum", "rewards", "stampHistory",
            "lastStampDate", "lastStampAt", "stampAchievedDate"}


def test_特典同期_値を反映しGASと同じ形で返す(client):
    m = 会員を作る(stamp_count=2)
    答 = 書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-1001",
                     "stampCount": 5, "stampCardNum": 2,
                     "rewards": [{"id": "r1", "cardNum": 1, "rewardName": "A賞", "used": False}],
                     "stampHistory": [{"acquiredDate": "2026-04-01T10:00:00+09:00"}],
                     "lastStampAt": "2026-04-01T10:00:00+09:00"}).json()
    assert 答["status"] == "ok"
    assert set(答["rewardStatus"]) == 特典の項目
    assert 答["rewardStatus"]["stampCount"] == 5
    assert 答["rewardStatus"]["stampCardNum"] == 2
    assert 答["rewardStatus"]["rewards"][0]["id"] == "r1"
    assert 答["rewardStatus"]["lastStampDate"] == "2026-04-01"
    m = 読み直す(m)
    assert m.stamp_count == 5 and m.stamp_card_number == 2
    assert m.reward_history[0]["id"] == "r1"
    assert timezone.localtime(m.last_stamp_at) == 日本時間(2026, 4, 1, 10, 0)


def test_特典同期_送らなかった項目はいまの値のまま(client):
    m = 会員を作る(stamp_count=7, stamp_card_number=3, reward_history=[{"id": "r1"}],
                stamp_history=[{"acquiredDate": "x"}])
    答 = 書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-1001"}).json()
    assert 答["rewardStatus"]["stampCount"] == 7
    m = 読み直す(m)
    assert m.stamp_count == 7 and m.stamp_card_number == 3
    assert m.reward_history == [{"id": "r1"}] and m.stamp_history == [{"acquiredDate": "x"}]


@pytest.mark.parametrize("値, 期待", [(None, 7), ("", 7), ("abc", 7), (9, 9), (0, 0), (99, 10)])
def test_特典同期_読めない値で0に戻さない(client, 値, 期待):
    """段階Bの下ごしらえ.md で決めた、GAS と違えている点（Number(x)||0 で0に戻る事故を防ぐ）。"""
    m = 会員を作る(stamp_count=7)
    書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-1001", "stampCount": 値})
    assert 読み直す(m).stamp_count == 期待


def test_特典同期_いない会員には行を作らない(client):
    """GAS は ensureUserRowFromActivity_ で行を足すが、名前も電話も無い行になる。足さない。"""
    答 = 書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-9999", "stampCount": 3}).json()
    assert 答 == {"status": "error", "message": "会員情報が見つかりません"}
    assert Member.objects.count() == 0


def test_特典同期_会員IDが無ければ断る(client):
    答 = 書く(client, {"type": "syncUserRewardStatus"}).json()
    assert 答 == {"status": "error", "message": "会員IDが指定されていません"}


def test_特典同期_stampAchievedDateを受けて保存する(client):
    m = 会員を作る(stamp_count=10)
    答 = 書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-1001", "stampCount": 10,
                     "stampAchievedDate": "2026-04-01T10:00:00+09:00"}).json()
    assert parse_datetime(答["rewardStatus"]["stampAchievedDate"]) == 日本時間(2026, 4, 1, 10, 0)
    assert timezone.localtime(読み直す(m).stamp_achieved_at) == 日本時間(2026, 4, 1, 10, 0)


def test_特典を見る_いない会員はエラーではなく空の状態(client):
    """登録の途中でも呼ばれるため（GAS 8232行 getDefaultRewardStatus_）。"""
    答 = 読む(client, "getUserRewardStatus", {"memberId": "MYM-9999"}).json()
    assert 答 == {
        "status": "ok",
        "rewardStatus": {"stampCount": 0, "stampCardNum": 1, "rewards": [], "stampHistory": [],
                         "lastStampDate": "", "lastStampAt": "", "stampAchievedDate": ""},
        "surveyAnswered": False,
        "adminSetAt": "",
    }


def test_特典を見る_会員IDが無ければエラー(client):
    答 = 読む(client, "getUserRewardStatus", {}).json()
    assert 答 == {"status": "error", "message": "会員IDが必要です"}


def test_特典を見る_いまの姿と管理者が直した時刻を返す(client):
    会員を作る(stamp_count=6, stamp_card_number=0, reward_history=[{"id": "r1"}],
            last_stamp_at=日本時間(2026, 4, 1, 10, 0), stamp_achieved_at=日本時間(2026, 3, 1, 9, 0),
            survey_answered_at=timezone.now(), reward_admin_set_at=日本時間(2026, 4, 2, 12, 30))
    答 = 読む(client, "getUserRewardStatus", {"memberId": "MYM-1001"}).json()
    s = 答["rewardStatus"]
    assert set(s) == 特典の項目
    assert s["stampCount"] == 6
    assert s["stampCardNum"] == 1                 # 0 は 1 に直す（GAS sanitizeRewardStatus_）
    assert s["rewards"] == [{"id": "r1"}]
    assert s["lastStampDate"] == "2026-04-01"
    assert parse_datetime(s["lastStampAt"]) == 日本時間(2026, 4, 1, 10, 0)
    assert parse_datetime(s["stampAchievedDate"]) == 日本時間(2026, 3, 1, 9, 0)
    assert 答["surveyAnswered"] is True
    # **これが空だと、受付で入れたスタンプがアプリを開いた瞬間に消える**（2026年8月）
    assert parse_datetime(答["adminSetAt"]) == 日本時間(2026, 4, 2, 12, 30)


def test_特典を見る_読むだけで書かない(client):
    """CLAUDE.md「記録を読む通信で書き込む」を禁ずる。保留のお礼スタンプも、ここでは触らない。"""
    m = 会員を作る(stamp_count=3, survey_stamp_pending_at=timezone.now())
    前 = 読み直す(m).changed_at
    読む(client, "getUserRewardStatus", {"memberId": "MYM-1001"})
    m = 読み直す(m)
    assert m.stamp_count == 3 and m.survey_stamp_pending_at is not None
    assert m.changed_at == 前


def test_保留のお礼スタンプは空きができたときに書く口で回収される(client):
    m = 会員を作る(stamp_count=10, survey_stamp_pending_at=timezone.now(), survey_answered_at=timezone.now())
    # 新しいカードに進んだ（アプリが 0 個・2枚目を同期してくる）
    書く(client, {"type": "syncUserRewardStatus", "memberId": "MYM-1001", "stampCount": 0, "stampCardNum": 2})
    m = 読み直す(m)
    assert m.stamp_count == 1
    assert m.survey_stamp_granted_at is not None
    assert m.survey_stamp_pending_at is None


# ═════════════════════════════════════════════════════════
# drawRewardGacha — 1枚のカードにつき1回（GAS 8039行）
# ═════════════════════════════════════════════════════════

def ガチャ(client, 会員ID="MYM-1001"):
    return 書く(client, {"type": "drawRewardGacha", "memberId": 会員ID}).json()


特典の見た目 = {"key", "rankLabel", "rewardName", "rewardNote", "earnedDate", "expiryDate",
            "capsuleColor", "accentColor", "message", "alreadyDrawn"}


def test_ガチャ_10個たまっていれば引ける(client):
    m = 会員を作る(stamp_count=10, stamp_card_number=1, stamp_achieved_at=日本時間(2026, 4, 1, 10, 0))
    答 = ガチャ(client)
    assert 答["status"] == "ok"
    assert "alreadyDrawn" not in 答 or 答["alreadyDrawn"] is False
    assert set(答["drawnReward"]) == 特典の見た目
    assert 答["drawnReward"]["alreadyDrawn"] is False
    assert 答["drawnReward"]["key"] in ("A", "B", "C", "D")
    assert 答["drawnReward"]["rewardName"].endswith("賞プレゼント")   # 設定が無ければ既定の名前
    assert set(答["rewardStatus"]) == 特典の項目
    m = 読み直す(m)
    assert len(m.reward_history) == 1
    r = m.reward_history[0]
    assert set(r) == {"id", "cardNum", "rewardName", "rewardNote", "earnedDate", "expiryDate", "used"}
    assert r["cardNum"] == 1 and r["used"] is False and r["id"].startswith("reward-")
    assert parse_datetime(r["earnedDate"]) == 日本時間(2026, 4, 1, 10, 0)


def test_ガチャ_同じカードでは2回引けない(client):
    m = 会員を作る(stamp_count=10, stamp_card_number=1)
    一回目 = ガチャ(client)
    二回目 = ガチャ(client)
    assert 二回目["status"] == "ok"
    assert 二回目["alreadyDrawn"] is True
    assert 二回目["drawnReward"]["alreadyDrawn"] is True
    assert 二回目["drawnReward"]["rewardName"] == 一回目["drawnReward"]["rewardName"]
    assert len(読み直す(m).reward_history) == 1


def test_ガチャ_カードが進めばまた引ける(client):
    m = 会員を作る(stamp_count=10, stamp_card_number=1)
    ガチャ(client)
    # 新しいカードへ進んだ（1枚目の特典は残したまま）
    Member.objects.filter(pk=m.pk).update(stamp_card_number=2, stamp_count=10)
    答 = ガチャ(client)
    assert 答.get("alreadyDrawn") is not True
    assert [r["cardNum"] for r in 読み直す(m).reward_history] == [2, 1]   # 新しいものが先頭


def test_ガチャ_10個未満は引けない(client):
    会員を作る(stamp_count=9)
    答 = ガチャ(client)
    assert 答 == {"status": "error", "message": "スタンプが10個たまっていません"}


def test_ガチャ_いない会員と会員IDなし(client):
    assert ガチャ(client, "MYM-9999") == {"status": "error", "message": "会員情報が見つかりません"}
    assert 書く(client, {"type": "drawRewardGacha"}).json() == {"status": "error", "message": "会員IDが指定されていません"}


def test_ガチャ_有効期限は達成日時の1か月後_JSのsetMonthと同じ(client):
    """1/31 → 2/31 は無いので 3/3（GAS buildRewardGachaEntry_ 4501行）。"""
    会員を作る(stamp_count=10, stamp_achieved_at=日本時間(2026, 1, 31, 8, 0))
    答 = ガチャ(client)
    assert parse_datetime(答["drawnReward"]["earnedDate"]) == 日本時間(2026, 1, 31, 8, 0)
    assert parse_datetime(答["drawnReward"]["expiryDate"]) == 日本時間(2026, 3, 3, 8, 0)


def test_ガチャ_有効期限は日本時間で月を進める(client):
    会員を作る(stamp_count=10, stamp_achieved_at=日本時間(2026, 5, 1, 5, 0))
    答 = ガチャ(client)
    assert parse_datetime(答["drawnReward"]["expiryDate"]) == 日本時間(2026, 6, 1, 5, 0)


def test_ガチャ_達成日時が無ければいまを使い記録する(client):
    m = 会員を作る(stamp_count=10, stamp_achieved_at=None)
    前 = timezone.now()
    答 = ガチャ(client)
    m = 読み直す(m)
    assert m.stamp_achieved_at is not None and m.stamp_achieved_at >= 前
    assert parse_datetime(答["rewardStatus"]["stampAchievedDate"]) == m.stamp_achieved_at.replace(microsecond=0)


def test_ガチャ_その月の景品表に従う(client):
    月 = timezone.localtime().strftime("%Y-%m")
    AppSetting.objects.create(key="REWARD_GACHA_CONFIG", value={"monthlyPrizes": [{
        "month": 月,
        "prizes": {"A": {"content": "お茶", "probability": 100, "note": "受付で"},
                   "B": {"probability": 0}, "C": {"probability": 0}, "D": {"probability": 0}},
    }]})
    会員を作る(stamp_count=10)
    答 = ガチャ(client)
    assert 答["drawnReward"]["key"] == "A"
    assert 答["drawnReward"]["rankLabel"] == "A賞"
    assert 答["drawnReward"]["rewardName"] == "A賞 お茶"
    assert 答["drawnReward"]["rewardNote"] == "受付で"
    assert 答["drawnReward"]["capsuleColor"] == "#f5cb6c"


# ═════════════════════════════════════════════════════════
# recoverAccount — 生年月日が必須（GAS 7685行）
# ═════════════════════════════════════════════════════════

会員の返す項目 = {"memberId", "name", "kana", "phone", "avatar", "memo", "status", "birthday",
              "address", "passcode", "deviceSessions", "stampCount", "stampCardNum", "rewards",
              "stampHistory", "lastStampDate", "lastStampAt", "stampAchievedAt", "regDate"}


def 復元(client, **項目):
    中 = {"type": "recoverAccount", "name": "佐藤花子", "birthday": "1990-04-05", "newPasscode": "5678"}
    中.update(項目)
    return 書く(client, 中).json()


def test_復元_氏名と生年月日で通る(client):
    m = 会員を作る(name="佐藤花子", kana="さとうはなこ", birthday="1990-04-05", phone="09012345678",
                passcode="1234", reward_history=[{"id": "r1"}], stamp_count=3,
                created_at=日本時間(2026, 1, 15, 10, 0))
    答 = 復元(client)
    assert 答["status"] == "ok"
    assert 答["recoveredBy"] == "identity"
    u = 答["user"]
    assert set(u) == 会員の返す項目
    assert u["memberId"] == "MYM-1001" and u["name"] == "佐藤花子"
    assert u["birthday"] == "1990-04-05"
    assert u["passcode"] == "5678"
    # rewards と stampHistory は **文字列**（GAS がセルの中身をそのまま返すため。app.js 7160行が JSON.parse する）
    assert isinstance(u["rewards"], str) and json.loads(u["rewards"]) == [{"id": "r1"}]
    assert isinstance(u["stampHistory"], str)
    assert u["stampCount"] == 3 and u["stampCardNum"] == 1
    assert u["regDate"] == "2026-01-15"
    m = 読み直す(m)
    assert check_password("5678", m.passcode_hash)
    assert m.registration_source == "復元" and m.registration_source_detail == "本人情報で復元"


def test_復元_フリガナと生年月日でも通る(client):
    会員を作る(name="佐藤花子", kana="サトウハナコ", birthday="1990-04-05")
    答 = 復元(client, name="", kana="さとう　はなこ")
    assert 答["status"] == "ok"


def test_復元_生年月日が違えば通らない(client):
    会員を作る(name="佐藤花子", birthday="1990-04-05")
    答 = 復元(client, birthday="1990-04-06")
    assert 答 == {"status": "error", "message": "一致する会員情報が見つかりませんでした。入力内容をご確認ください。"}


@pytest.mark.parametrize("項目, 文言", [
    ({"name": "", "kana": ""}, "お名前またはフリガナを入力してください。"),
    ({"birthday": ""}, "生年月日を入力してください。"),
    ({"newPasscode": "123"}, "この端末で使うパスコードを4桁または6桁の数字で入力してください。"),
])
def test_復元_足りない入力の文言はGASと同じ(client, 項目, 文言):
    会員を作る(name="佐藤花子", birthday="1990-04-05")
    assert 復元(client, **項目) == {"status": "error", "message": 文言}


def test_復元_電話番号を入れたら絞り込みに使う(client):
    会員を作る("MYM-1001", name="佐藤花子", birthday="1990-04-05", phone="090-1234-5678")
    会員を作る("MYM-1002", name="佐藤花子", birthday="1990-04-05", phone="080-0000-0000")
    答 = 復元(client)
    assert "複数見つかりました" in 答["message"]
    答 = 復元(client, phone="9012345678")       # 先頭の0が落ちていても一致
    assert 答["user"]["memberId"] == "MYM-1001"


def test_復元_いまのパスコードでも絞り込める(client):
    会員を作る("MYM-1001", name="佐藤花子", birthday="1990-04-05", passcode="1111")
    会員を作る("MYM-1002", name="佐藤花子", birthday="1990-04-05", passcode="2222")
    assert 復元(client, passcode="2222")["user"]["memberId"] == "MYM-1002"


def test_復元_退会の印がある方も戻れて印が消える(client):
    m = 会員を作る(name="佐藤花子", birthday="1990-04-05", deleted=True,
                deleted_at=timezone.now(), merged_into_id="MYM-2000")
    assert 復元(client)["status"] == "ok"
    m = 読み直す(m)
    assert m.deleted is False and m.deleted_at is None and m.merged_into_id == ""


def test_復元_10回はずすとお休み_複数一致は数えない(client):
    会員を作る(name="佐藤花子", birthday="1990-04-05")
    for _ in range(10):
        復元(client, birthday="2000-01-01")
    答 = 復元(client)
    assert 答["message"] == "復元の試行が続いたため、一時的に停止しています。10分後に再度お試しください。"


def test_復元_引き継ぎコードで通る_コードは1回で消える(client):
    m = 会員を作る(name="佐藤花子", transfer_code="12345678", transfer_code_issued_at=timezone.now())
    答 = 書く(client, {"type": "recoverAccount", "transferCode": "1234-5678", "newPasscode": "5678"}).json()
    assert 答["status"] == "ok" and 答["recoveredBy"] == "transferCode"
    assert 答["user"]["name"] == "佐藤花子"      # 入口はここからお名前を受け取る
    m = 読み直す(m)
    assert m.transfer_code == "" and m.transfer_code_issued_at is None
    assert m.registration_source == "引き継ぎコード利用"
    assert check_password("5678", m.passcode_hash)


def test_復元_引き継ぎコードは数え上げの対象外(client):
    会員を作る(name="佐藤花子", birthday="1990-04-05", transfer_code="12345678",
            transfer_code_issued_at=timezone.now())
    for _ in range(10):
        復元(client, birthday="2000-01-01")
    答 = 書く(client, {"type": "recoverAccount", "transferCode": "12345678", "newPasscode": "5678"}).json()
    assert 答["status"] == "ok"


def test_復元_発行日時の無いコードは無効でその場で消す(client):
    m = 会員を作る(transfer_code="12345678", transfer_code_issued_at=None)
    答 = 書く(client, {"type": "recoverAccount", "transferCode": "12345678", "newPasscode": "5678"}).json()
    assert 答["message"] == "この引き継ぎコードは無効です。元の端末で新しいコードを発行してください。"
    assert 読み直す(m).transfer_code == ""


def test_復元_1週間を過ぎたコードは期限切れ(client):
    m = 会員を作る(transfer_code="12345678",
                transfer_code_issued_at=timezone.now() - dt.timedelta(hours=169))
    答 = 書く(client, {"type": "recoverAccount", "transferCode": "12345678", "newPasscode": "5678"}).json()
    assert 答["message"] == "この引き継ぎコードは期限切れです。元の端末で新しいコードを発行してください。"
    assert 読み直す(m).transfer_code == ""


def test_復元_知らないコードは見つからない(client):
    答 = 書く(client, {"type": "recoverAccount", "transferCode": "00000000", "newPasscode": "5678"}).json()
    assert 答["message"] == "一致する会員情報が見つかりませんでした。入力内容をご確認ください。"


# ═════════════════════════════════════════════════════════
# resetForgottenPasscode — 電話か生年月日（GAS 7862行）
# ═════════════════════════════════════════════════════════

def 再設定(client, **項目):
    中 = {"type": "resetForgottenPasscode", "name": "佐藤花子", "newPasscode": "5678"}
    中.update(項目)
    return 書く(client, 中).json()


def test_再設定_記録に電話だけなら電話で通る(client):
    m = 会員を作る(name="佐藤花子", phone="09012345678", birthday=None, passcode="1234")
    答 = 再設定(client, phone="090-1234-5678")
    assert 答["status"] == "ok"
    assert set(答["user"]) == 会員の返す項目 and 答["user"]["passcode"] == "5678"
    assert check_password("5678", 読み直す(m).passcode_hash)
    記録 = AuditLog.objects.get(kind="パスコード再設定")
    assert 記録.result == "成功" and 記録.target == "MYM-1001" and 記録.operator == "ご本人"
    assert 記録.summary == "氏名と、記録にある1項目で確認（もう一方は未登録）"
    assert 記録.sheet_row is None


def test_再設定_記録に生年月日だけなら生年月日で通る(client):
    会員を作る(name="佐藤花子", phone="", birthday="1990-04-05")
    assert 再設定(client, birthday="1990/04/05")["status"] == "ok"


def test_再設定_記録に両方あれば両方要る(client):
    会員を作る(name="佐藤花子", phone="09012345678", birthday="1990-04-05")
    assert 再設定(client, phone="09012345678")["status"] == "error"
    assert 再設定(client, birthday="1990-04-05")["status"] == "error"
    答 = 再設定(client, phone="09012345678", birthday="1990-04-05")
    assert 答["status"] == "ok"
    assert AuditLog.objects.get(kind="パスコード再設定").summary == "氏名・電話番号・生年月日の3点で確認"


def test_再設定_記録に電話も生年月日も無い方は通れない(client):
    """お名前だけで通すと同姓同名の他人が入れる。受付でご対応。"""
    会員を作る(name="佐藤花子", phone="", birthday=None)
    答 = 再設定(client, phone="09012345678", birthday="1990-04-05")
    assert 答["status"] == "error"
    assert 答["message"].startswith("ご登録の内容と一致しませんでした。")
    assert "受付にお申し出ください" in 答["message"]


def test_再設定_足りない入力の文言はGASと同じ(client):
    assert 再設定(client)["message"] == "お名前と、電話番号または生年月日のどちらか、そして新しいパスコードを入力してください。"
    assert 再設定(client, phone="09012345678", newPasscode="12")["message"] == "新しいパスコードは4桁または6桁の数字で入力してください。"


def test_再設定_複数一致は会員IDで絞る(client):
    会員を作る("MYM-1001", name="佐藤花子", phone="09012345678")
    会員を作る("MYM-1002", name="佐藤花子", phone="09012345678")
    答 = 再設定(client, phone="09012345678")
    assert 答["message"] == "同じ内容のご登録が複数見つかりました。恐れ入りますが、受付にお申し出ください。"
    答 = 再設定(client, phone="09012345678", memberId="MYM-1002")
    assert 答["user"]["memberId"] == "MYM-1002"


def test_再設定_退会の印が消え引き継ぎコードも消える(client):
    m = 会員を作る(name="佐藤花子", phone="09012345678", deleted=True,
                transfer_code="12345678", transfer_code_issued_at=timezone.now())
    再設定(client, phone="09012345678")
    m = 読み直す(m)
    assert m.deleted is False and m.transfer_code == ""


def test_再設定_10回はずすとお休み(client):
    会員を作る(name="佐藤花子", phone="09012345678")
    for _ in range(10):
        再設定(client, phone="00000000000")
    答 = 再設定(client, phone="09012345678")
    assert 答["message"] == "お手続きが続いたため、一時的にお休みしています。10分ほど経ってから、もう一度お試しください。"


# ═════════════════════════════════════════════════════════
# getRecoveryCandidates — 合鍵なしで呼べるので、返すものを間違えると漏れる（GAS 7623行）
# ═════════════════════════════════════════════════════════

def 候補(client, **引数):
    return 読む(client, "getRecoveryCandidates", 引数).json()


def test_復元候補_返すのは会員IDと伏せた名前と理由だけ(client):
    会員を作る(name="佐藤花子", kana="さとうはなこ", phone="09012345678", birthday="1990-04-05")
    答 = 候補(client, name="佐藤花子")
    assert 答["status"] == "ok"
    assert 答["candidates"] == [{"memberId": "MYM-1001", "name": "佐○○○", "reasons": ["氏名"]}]
    文 = json.dumps(答, ensure_ascii=False)
    for 秘密 in ("1990", "0901234", "さとうはなこ", "佐藤花子"):
        assert 秘密 not in 文


def test_復元候補_お名前の伏せ方はGASのmaskNameForRecoveryと同じ(client):
    for 名, 期待 in [("佐藤", "佐○"), ("佐藤花", "佐○○"), ("佐藤花子", "佐○○○"),
                   ("佐藤花子太郎", "佐○○○"), ("佐", "佐○")]:
        assert views._お名前を伏せる(名) == 期待
    assert views._お名前を伏せる("") == "会員"


def test_復元候補_生年月日と電話は3点で氏名より上に並ぶ(client):
    会員を作る("MYM-1001", name="佐藤花子", birthday="1980-01-01")
    会員を作る("MYM-1002", name="山田太郎", birthday="1990-04-05", phone="09012345678")
    答 = 候補(client, name="佐藤花子", birthday="1990-04-05", phone="9012345678")
    assert [(c["memberId"], c["reasons"]) for c in 答["candidates"]] == [
        ("MYM-1002", ["電話番号", "生年月日"]),
        ("MYM-1001", ["氏名"]),
    ]


def test_復元候補_上位5件まで(client):
    for i in range(7):
        会員を作る(f"MYM-10{i:02d}", name="佐藤花子")
    assert len(候補(client, name="佐藤花子")["candidates"]) == 5


def test_復元候補_何も無ければ空_退会者は出ない(client):
    会員を作る(name="佐藤花子", deleted=True)
    assert 候補(client)["candidates"] == []
    assert 候補(client, name="佐藤花子")["candidates"] == []
    assert 候補(client, name="だれか")["candidates"] == []


def test_復元候補_フリガナはカタカナでも一致(client):
    会員を作る(name="佐藤花子", kana="サトウハナコ")
    assert 候補(client, kana="さとう はなこ")["candidates"][0]["reasons"] == ["フリガナ"]


# ═════════════════════════════════════════════════════════
# unsubscribePush — 会員の行を空にする部分（GAS 8136行）
# ═════════════════════════════════════════════════════════

def test_通知を切る_届け先が空になり他の項目は残る(client):
    m = 会員を作る(push_subscription="sub-abc", push_enabled=True, memo="覚え書き", stamp_count=4)
    答 = 書く(client, {"type": "unsubscribePush", "memberId": "MYM-1001", "playerId": "p-1"}).json()
    assert 答["status"] == "ok"
    m = 読み直す(m)
    assert m.push_subscription == "" and m.push_enabled is False
    assert m.memo == "覚え書き" and m.stamp_count == 4
    # 届け先の一覧からも消える
    assert 読む(client, "getPushUsers", 合鍵あり=True).json()["users"] == []


def test_通知を切る_いない会員と会員IDなし(client):
    assert 書く(client, {"type": "unsubscribePush", "memberId": "MYM-9999"}).json()["status"] == "error"
    assert 書く(client, {"type": "unsubscribePush"}).json()["status"] == "error"
