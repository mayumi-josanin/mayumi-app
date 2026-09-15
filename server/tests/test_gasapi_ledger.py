"""予約システム → 台帳の窓口（upsertLedgerFromReservation）と、
新規登録（registerAccount）の LINEユーザーID 照合。

設計: docs/design/予約のお客様を台帳へ入れる.md ①②③、予約からLINEIDを集める.md。
"""

import pytest

from apps.gasapi import entrance
from apps.manage import member_gate
from apps.members.models import Member
from tests.gasapi_member_support import 会員を作る, 書く, 読み直す

pytestmark = [pytest.mark.django_db, pytest.mark.usefixtures("会員試験の下ごしらえ")]


@pytest.fixture
def サーバーが正():
    member_gate.切り替える("server")


def 送る(client, 合鍵あり=True, **項目):
    中 = {"type": "upsertLedgerFromReservation", "name": "山田 花子", "kana": "やまだ はなこ",
          "phone": "090-1234-5678", "birthday": "1990-04-05", "address": "平塚市1-2-3",
          "lineUserId": "U-line-0001", "reservationId": "42"}
    中.update(項目)
    return 書く(client, 中, 合鍵あり=合鍵あり)


# ═════════════════════════════════════════════════════════
# 門
# ═════════════════════════════════════════════════════════

def test_合鍵なしは403(client, サーバーが正):
    """お客様アプリの窓口ではない。公開アクションに入れない。"""
    r = 送る(client, 合鍵あり=False)
    assert r.status_code == 403
    assert Member.objects.count() == 0


def test_会員の正がシートのうちは書かずにretryLater(client):
    答 = 送る(client).json()
    assert 答["status"] == "error"
    assert 答["retryLater"] is True
    assert "9/19" in 答["message"]
    assert Member.objects.count() == 0


def test_お名前が空なら断る(client, サーバーが正):
    答 = 送る(client, name="  ").json()
    assert 答["status"] == "error"
    assert Member.objects.count() == 0


# ═════════════════════════════════════════════════════════
# 4つの照合
# ═════════════════════════════════════════════════════════

def test_LINEユーザーIDで結ぶ(client, サーバーが正):
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", line_user_id="U-line-0001",
                birthday=None, phone="")
    答 = 送る(client, name="山田　花子", birthday="1980-01-01", phone="09099999999").json()
    assert 答["status"] == "ok"
    assert 答["result"] == "linked"
    assert 答["memberId"] == "MYM-1001"
    assert 答["matchedBy"] == "line"
    assert Member.objects.count() == 1
    m = 読み直す(m)
    # 空いていた欄が埋まる
    assert str(m.birthday) == "1980-01-01"
    assert m.phone == "09099999999"


def test_氏名と生年月日で結ぶ(client, サーバーが正):
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05")
    答 = 送る(client).json()
    assert 答["result"] == "linked"
    assert 答["matchedBy"] == "nameBirthday"
    assert 答["memberId"] == "MYM-1001"
    assert 読み直す(m).line_user_id == "U-line-0001"
    assert "lineUserId" in 答["filled"]


def test_氏名と生年月日は空白とスラッシュの違いを無視する(client, サーバーが正):
    会員を作る("MYM-1001", name="山田 花子", passcode="", birthday="1990-04-05")
    答 = 送る(client, name="山田　花子", birthday="1990/04/05").json()
    assert 答["result"] == "linked" and 答["matchedBy"] == "nameBirthday"


def test_氏名と電話で結ぶ(client, サーバーが正):
    """生年月日が台帳側に無い方。電話はハイフンと先頭の0の違いを揃えて比べる。"""
    会員を作る("MYM-1001", name="山田花子", passcode="", birthday=None, phone="9012345678")
    答 = 送る(client, phone="090-1234-5678").json()
    assert 答["result"] == "linked"
    assert 答["matchedBy"] == "namePhone"
    assert 答["memberId"] == "MYM-1001"


def test_生年月日が違えば電話では結ばない_同姓同名の別人(client, サーバーが正):
    """氏名＋生年月日で当たらなくても、電話が同じなら結ぶ（家族で1台）。
    ここは設計書の順のとおり。ただし生年月日は空欄にしか書かないので、
    台帳の生年月日が入っている方には別の日付は入らない。"""
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1970-01-01", phone="09012345678")
    答 = 送る(client, birthday="1990-04-05").json()
    assert 答["matchedBy"] == "namePhone"
    assert str(読み直す(m).birthday) == "1970-01-01"


def test_どれも無ければ新しく作る(client, サーバーが正, monkeypatch):
    monkeypatch.setattr(entrance.random, "randrange", lambda *a, **k: 4321)
    答 = 送る(client, note="初回はビジリス希望").json()
    assert 答["status"] == "ok"
    assert 答["result"] == "created"
    assert 答["memberId"] == "MYM-4321"
    assert 答["matchedBy"] is None
    m = Member.objects.get(pk="MYM-4321")
    assert m.name == "山田花子"                 # 空白は取って保存（新規登録と同じ）
    assert m.kana == "やまだはなこ"
    assert m.phone == "090-1234-5678"
    assert str(m.birthday) == "1990-04-05"
    assert m.address == "平塚市1-2-3"
    assert m.line_user_id == "U-line-0001"
    assert m.registration_source == "予約システム"
    assert m.registration_source_detail == "42"
    assert m.registration_source_updated_at is not None
    assert m.memo.startswith("予約システムから取り込み（20")
    assert m.memo.endswith("）\n初回はビジリス希望")
    # パスコードは空。まだアプリには入れない
    assert m.passcode_hash == ""
    assert set(答["filled"]) == {"kana", "phone", "birthday", "address", "lineUserId"}


def test_採番は既存と重ならない(client, サーバーが正, monkeypatch):
    会員を作る("MYM-1001", name="別の方")
    出す = iter([1001, 1001, 2002])
    monkeypatch.setattr(entrance.random, "randrange", lambda *a, **k: next(出す))
    assert 送る(client).json()["memberId"] == "MYM-2002"


def test_採番できなければ作らずに断る(client, サーバーが正, monkeypatch):
    会員を作る("MYM-1001", name="別の方")
    monkeypatch.setattr(entrance.random, "randrange", lambda *a, **k: 1001)
    答 = 送る(client).json()
    assert 答["status"] == "error"
    assert "採番" in 答["message"]
    assert Member.objects.count() == 1


def test_退会者は候補にしない(client, サーバーが正):
    会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05", deleted=True)
    答 = 送る(client, lineUserId="").json()
    assert 答["result"] == "created"
    assert 答["memberId"] != "MYM-1001"


# ═════════════════════════════════════════════════════════
# 保留（結ばない・作らない）
# ═════════════════════════════════════════════════════════

def test_氏名と生年月日が2人当たれば保留(client, サーバーが正):
    会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05")
    会員を作る("MYM-1002", name="山田花子", passcode="", birthday="1990-04-05")
    答 = 送る(client).json()
    assert 答["status"] == "ok"
    assert 答["result"] == "hold"
    assert 答["candidates"] == 2
    assert 答["memberId"] is None
    assert Member.objects.count() == 2
    assert not Member.objects.filter(line_user_id="U-line-0001").exists()


def test_氏名と電話が2人当たれば保留(client, サーバーが正):
    会員を作る("MYM-1001", name="山田花子", passcode="", phone="09012345678")
    会員を作る("MYM-1002", name="山田花子", passcode="", phone="090-1234-5678")
    答 = 送る(client, birthday="").json()
    assert 答["result"] == "hold" and 答["candidates"] == 2
    assert Member.objects.count() == 2


def test_退会した行がそのLINEIDを持っていれば保留(client, サーバーが正):
    """LINEユーザーID は unique。退会者の行に残っていると書けないので、受付が見る。"""
    会員を作る("MYM-1001", name="山田花子", passcode="", line_user_id="U-line-0001", deleted=True)
    会員を作る("MYM-1002", name="山田花子", passcode="", birthday="1990-04-05")
    答 = 送る(client).json()
    assert 答["result"] == "hold"
    assert 答["reason"] == "lineTaken"
    assert 読み直す(Member.objects.get(pk="MYM-1002")).line_user_id is None


# ═════════════════════════════════════════════════════════
# 空いている欄だけ埋める
# ═════════════════════════════════════════════════════════

def test_結んだ行の入っている欄は上書きしない(client, サーバーが正):
    m = 会員を作る("MYM-1001", name="山田花子", passcode="1234", birthday="1990-04-05",
                kana="やまだはなこ", phone="0463111111", address="寒川町", memo="受付のメモ")
    答 = 送る(client, kana="別のかな", phone="09099999999", address="別の住所").json()
    assert 答["result"] == "linked"
    assert 答["filled"] == ["lineUserId"]
    m = 読み直す(m)
    assert m.kana == "やまだはなこ"
    assert m.phone == "0463111111"
    assert m.address == "寒川町"
    assert m.memo == "受付のメモ"          # 結んだ行のメモは触らない
    assert m.passcode_hash                # ログインの設定も触らない


def test_結んだ行の空いている欄は埋まる(client, サーバーが正):
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05",
                kana="", phone="", address="")
    答 = 送る(client).json()
    assert set(答["filled"]) == {"kana", "phone", "address", "lineUserId"}
    m = 読み直す(m)
    assert m.kana == "やまだはなこ" and m.phone == "090-1234-5678" and m.address == "平塚市1-2-3"


# ═════════════════════════════════════════════════════════
# LINEユーザーIDの不一致
# ═════════════════════════════════════════════════════════

def test_LINEIDが入っていて違う値なら上書きしない(client, サーバーが正):
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05",
                line_user_id="U-line-old")
    答 = 送る(client, lineUserId="U-line-new").json()
    assert 答["result"] == "linked"
    assert 答["matchedBy"] == "nameBirthday"
    assert 答["lineMismatch"] is True
    assert 読み直す(m).line_user_id == "U-line-old"


def test_LINEIDが別の名前の行に付いていれば_そのIDは使わない_代理のご予約(client, サーバーが正):
    """お母様の LINE で娘さんのご予約。ID だけで結ぶと、お母様の行に娘さんの
    生年月日や電話が入る。お名前が違えば ID は無かったことにして、娘さんの行を探す・作る。"""
    母 = 会員を作る("MYM-1001", name="山田良子", passcode="", line_user_id="U-line-0001",
                 birthday="1960-01-01", phone="")
    答 = 送る(client, name="山田花子").json()
    assert 答["result"] == "created"
    assert 答["memberId"] != "MYM-1001"
    assert 答["lineMismatch"] is True
    娘 = Member.objects.get(pk=答["memberId"])
    assert 娘.line_user_id is None                      # お母様の ID を娘さんに付けない
    母 = 読み直す(母)
    assert 母.phone == "" and str(母.birthday) == "1960-01-01"   # お母様の行は触らない


def test_LINEIDが別の名前の行に付いていても_娘さんの既存の行には結ぶ(client, サーバーが正):
    会員を作る("MYM-1001", name="山田良子", passcode="", line_user_id="U-line-0001")
    会員を作る("MYM-1002", name="山田花子", passcode="", birthday="1990-04-05")
    答 = 送る(client, name="山田花子").json()
    assert 答["result"] == "linked" and 答["memberId"] == "MYM-1002"
    assert 答["lineMismatch"] is True
    assert 読み直す(Member.objects.get(pk="MYM-1002")).line_user_id is None


def test_同じ内容を2回送っても行は増えない(client, サーバーが正):
    """通信は失敗して再送される。2回届いても壊れない。"""
    一 = 送る(client).json()
    二 = 送る(client).json()
    assert 一["result"] == "created"
    assert 二["result"] == "linked" and 二["matchedBy"] == "line"
    assert 二["memberId"] == 一["memberId"]
    assert Member.objects.count() == 1


# ═════════════════════════════════════════════════════════
# registerAccount の LINE 照合（③）
# ═════════════════════════════════════════════════════════

def 登録(client, **項目):
    中 = {"type": "registerAccount", "name": "山田花子", "kana": "やまだはなこ",
          "birthday": "1990-04-05", "phone": "09012345678", "address": "平塚市1-2-3",
          "passcode": "1234"}
    中.update(項目)
    return 書く(client, 中).json()


def test_新規登録_LINEIDで予約から作られた行を引き受ける(client):
    """予約から入った行（生年月日が違う書き方・空でも）に、アプリからパスコードを付ける。"""
    m = 会員を作る("MYM-1001", name="山田花子", passcode="", line_user_id="U-line-0001",
                birthday=None, phone="", kana="", memo="予約システムから取り込み（2026-09-24）")
    答 = 登録(client, lineUserId="U-line-0001")
    assert 答["status"] == "ok"
    assert 答["memberId"] == "MYM-1001"
    assert 答["claimed"] is True
    assert Member.objects.count() == 1
    m = 読み直す(m)
    assert m.passcode_hash
    assert str(m.birthday) == "1990-04-05"       # 空だった生年月日も埋まる
    assert m.kana == "やまだはなこ" and m.phone == "09012345678"
    assert m.memo == "予約システムから取り込み（2026-09-24）"


def test_新規登録_LINEIDが一致してもお名前が違えば引き受けない(client):
    """お母様の LINE で予約した娘さんが登録するとき。お母様の行を乗っ取らない。"""
    会員を作る("MYM-1001", name="山田良子", passcode="", line_user_id="U-line-0001",
             birthday="1960-01-01")
    答 = 登録(client, name="山田花子", lineUserId="U-line-0001")
    assert 答["status"] == "ok"
    assert 答["memberId"] != "MYM-1001"
    assert Member.objects.count() == 2
    assert 読み直す(Member.objects.get(pk="MYM-1001")).passcode_hash == ""


def test_新規登録_LINEIDで当たった行にパスコードがあればログインへ(client):
    会員を作る("MYM-1001", name="山田花子", passcode="1234", line_user_id="U-line-0001")
    答 = 登録(client, lineUserId="U-line-0001", birthday="1999-09-09")
    assert 答["status"] == "error"
    assert "ログイン" in 答["message"]


def test_新規登録_LINEIDが無ければこれまでどおり氏名と生年月日(client):
    会員を作る("MYM-1001", name="山田花子", passcode="", birthday="1990-04-05")
    答 = 登録(client)
    assert 答["memberId"] == "MYM-1001"
