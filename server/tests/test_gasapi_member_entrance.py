"""入口の3窓口。**ここが止まるとお客様がアプリに入れなくなる。**

    loginAccount        管理者・お客様.js 2536行
    registerAccount     同 2363行
    registerBijirisUse  同 2594行

GAS は `SERVER_TABLES` に member が入ると、この3つを **サーバーの答えのまま**
入口（start/index.html）へ返す（`_会員をサーバーへ_` 5215行は答えを加工しない）。
だから、サーバーの答えは GAS の `buildAccountSession_`（2301行）と同じ形で
なければならない。入口は `json.token` が無いと「ログインできませんでした」にする
（start/index.html 677行・749行）。
"""

import base64
import hashlib

import pytest
from django.contrib.auth.hashers import check_password

from apps.gasapi import entrance
from apps.members.models import Member
from tests.gasapi_member_support import 会員を作る, 書く, 読み直す

pytestmark = [pytest.mark.django_db, pytest.mark.usefixtures("会員試験の下ごしらえ")]


def ログイン(client, name, passcode):
    return 書く(client, {"type": "loginAccount", "name": name, "passcode": passcode}).json()


def 登録(client, **項目):
    中 = {"type": "registerAccount", "name": "鈴木一郎", "kana": "すずきいちろう",
          "birthday": "1990-04-05", "phone": "09012345678", "address": "平塚市1-2-3",
          "passcode": "1234"}
    中.update(項目)
    return 書く(client, 中).json()


# ═════════════════════════════════════════════════════════
# loginAccount
# ═════════════════════════════════════════════════════════

def test_ログイン_正しいパスコードで通る(client):
    m = 会員を作る(name="佐藤花子", kana="さとうはなこ", passcode="1234")
    答 = ログイン(client, "佐藤花子", "1234")
    assert 答["status"] == "ok"
    assert 答["memberId"] == "MYM-1001"
    assert 答["name"] == "佐藤花子"
    assert 答["kana"] == "さとうはなこ"
    # アプリを開いた記録（GAS の touchLastOnline_）
    assert 読み直す(m).last_online_at is not None


def test_ログイン_お名前の空白と全角半角は無視する(client):
    """GAS の normalizeNameForMatch_（7368行）。「山田　太郎」でも「山田太郎」に入れる。"""
    会員を作る(name="山田太郎", passcode="1234")
    assert ログイン(client, "山田　太郎", "1234")["status"] == "ok"
    assert ログイン(client, " 山田 太郎 ", "1234")["status"] == "ok"


def test_ログイン_password_の呼び名でも受ける(client):
    """古い呼び出しのために password も受ける（GAS 2540行）。"""
    会員を作る(name="佐藤花子", passcode="1234")
    答 = 書く(client, {"type": "loginAccount", "name": "佐藤花子", "password": "1234"}).json()
    assert 答["status"] == "ok"


def test_ログイン_違うパスコードは通らない(client):
    会員を作る(name="佐藤花子", passcode="1234")
    答 = ログイン(client, "佐藤花子", "9999")
    assert 答 == {"status": "error", "message": "お名前またはパスコードが違います。"}


def test_ログイン_いないお名前も同じ文言で断る(client):
    """存在を推測させない（GAS 2574行のコメント）。"""
    答 = ログイン(client, "だれか", "1234")
    assert 答["message"] == "お名前またはパスコードが違います。"


def test_ログイン_お名前かパスコードが空なら入力を促す(client):
    assert ログイン(client, "", "1234")["message"] == "お名前とパスコードを入力してください。"
    assert ログイン(client, "佐藤花子", "")["message"] == "お名前とパスコードを入力してください。"


def test_ログイン_同姓同名で同じパスコードは見分けがつかないので通さない(client):
    会員を作る("MYM-1001", name="佐藤花子", passcode="1234")
    会員を作る("MYM-1002", name="佐藤花子", passcode="1234")
    答 = ログイン(client, "佐藤花子", "1234")
    assert 答["status"] == "error"
    assert 答["message"] == "お名前またはパスコードが違います。"


def test_ログイン_同姓同名でもパスコードが違えば通る(client):
    会員を作る("MYM-1001", name="佐藤花子", passcode="1234")
    会員を作る("MYM-1002", name="佐藤花子", passcode="5678")
    assert ログイン(client, "佐藤花子", "5678")["memberId"] == "MYM-1002"


def test_ログイン_退会の印がある方は通らない(client):
    会員を作る(name="佐藤花子", passcode="1234", deleted=True)
    assert ログイン(client, "佐藤花子", "1234")["status"] == "error"


def test_ログイン_10回はずすとお休み(client):
    """GAS の ADMIN_LOGIN_MAX_ATTEMPTS = 10（1941行）。正しいパスコードでも止まる。"""
    会員を作る(name="佐藤花子", passcode="1234")
    for _ in range(10):
        ログイン(client, "佐藤花子", "0000")
    答 = ログイン(client, "佐藤花子", "1234")
    assert 答["status"] == "error"
    assert 答["message"] == "ログイン失敗が続いたため、一時的に停止しています。10分後に再試行してください。"


def test_ログイン_お休みはお名前ごと_他の方は入れる(client):
    """全体で数えると、攻撃者が1人いるだけで全員が締め出される。"""
    会員を作る("MYM-1001", name="佐藤花子", passcode="1234")
    会員を作る("MYM-1002", name="山田太郎", passcode="5678")
    for _ in range(10):
        ログイン(client, "佐藤花子", "0000")
    assert ログイン(client, "山田太郎", "5678")["status"] == "ok"


def test_ログイン_正しく入れば数え直しになる(client):
    会員を作る(name="佐藤花子", passcode="1234")
    for _ in range(9):
        ログイン(client, "佐藤花子", "0000")
    assert ログイン(client, "佐藤花子", "1234")["status"] == "ok"
    # ここで 0 に戻っているので、あと9回はずしても止まらない
    for _ in range(9):
        ログイン(client, "佐藤花子", "0000")
    assert ログイン(client, "佐藤花子", "1234")["status"] == "ok"


def test_ログイン_お名前の空白違いでもはずれは同じ相手として数える(client):
    """「佐藤 花子」と「佐藤花子」で別々に数えると、上限が2倍になる。"""
    会員を作る(name="佐藤花子", passcode="1234")
    for _ in range(5):
        ログイン(client, "佐藤花子", "0000")
    for _ in range(5):
        ログイン(client, "佐藤　花子", "0000")
    assert "停止" in ログイン(client, "佐藤花子", "1234")["message"]


@pytest.mark.xfail(reason="入口（start/index.html 677行）は json.token が無いと"
                          "「ログインできませんでした」にする。GAS の buildAccountSession_"
                          "（2301行）は role / bijiris / token / memberToken / expiresAt / "
                          "missing を返すが、サーバーの entrance.ログイン は memberId / name / kana "
                          "しか返さない。**member を渡した瞬間に誰もログインできなくなる**", strict=True)
def test_ログイン_入口が要る項目をそろえて返す(client):
    会員を作る(name="佐藤花子", kana="さとうはなこ", passcode="1234",
            phone="09012345678", birthday="1990-04-05", address="平塚市")
    答 = ログイン(client, "佐藤花子", "1234")
    assert 答["role"] == "member"
    assert 答["bijiris"] is False
    assert 答["token"] and 答["memberToken"]
    assert 答["expiresAt"]
    assert 答["missing"] == []


@pytest.mark.xfail(reason="GAS 2301行: role が「管理者」なら role:'admin' と管理者の札を返す。"
                          "サーバーは role を見ていない", strict=True)
def test_ログイン_管理者には管理者の札を返す(client):
    会員を作る(name="院長", passcode="1234", role="管理者")
    答 = ログイン(client, "院長", "1234")
    assert 答["role"] == "admin"
    assert 答["token"]


@pytest.mark.xfail(reason="GAS 2323行: 電話・生年月日・住所・フリガナが無い方には missing を返し、"
                          "入口がお願いする材料にする。サーバーは返さない", strict=True)
def test_ログイン_足りない項目をmissingで知らせる(client):
    会員を作る(name="佐藤花子", passcode="1234")
    答 = ログイン(client, "佐藤花子", "1234")
    assert set(答["missing"]) == {"kana", "phone", "birthday", "address"}


def _GASの旧パスワードハッシュ(password, salt, secret):
    """GAS の hashAccountPassword_（2207行）。salt|password|秘密 を SHA-256 → base64url。"""
    素 = (salt + "|" + password + "|" + secret).encode("utf-8")
    return base64.urlsafe_b64encode(hashlib.sha256(素).digest()).decode("ascii")


@pytest.mark.xfail(reason="旧方式の2名（password_hash + password_salt）の照合が GAS と違う。"
                          "GAS 2207行は sha256(salt|password|ADMIN_TOKEN_SECRET) の base64url、"
                          "entrance.py 148行は sha256(password+salt) の hex。"
                          "取り込んだハッシュでは一致せず、この2名は入れない", strict=True)
def test_ログイン_旧パスワード方式の方も入れる(client, settings):
    settings.ADMIN_TOKEN_SECRET = "secret-for-test"
    会員を作る(name="旧会員", passcode="", password_salt="salt-1",
            password_hash=_GASの旧パスワードハッシュ("abcd1234", "salt-1", "secret-for-test"))
    assert ログイン(client, "旧会員", "abcd1234")["status"] == "ok"


# ═════════════════════════════════════════════════════════
# registerAccount
# ═════════════════════════════════════════════════════════

def test_新規登録_会員IDはMYMと4桁で採番される(client):
    答 = 登録(client)
    assert 答["status"] == "ok"
    assert answer_is_new_member_id(答["memberId"])
    m = Member.objects.get(pk=答["memberId"])
    assert m.name == "鈴木一郎"
    assert m.kana == "すずきいちろう"
    assert m.phone == "09012345678"
    assert str(m.birthday) == "1990-04-05"
    assert m.address == "平塚市1-2-3"
    # **平文で持たない。**
    assert m.passcode_hash and m.passcode_hash != "1234"
    assert check_password("1234", m.passcode_hash)
    assert m.created_at is not None
    assert m.device_sessions == []


def answer_is_new_member_id(値):
    import re

    return re.fullmatch(r"MYM-\d{4}", 値) is not None and 1000 <= int(値[4:]) <= 9999


def test_新規登録_採番は既存と重ならない(client, monkeypatch):
    """乱数が使用中の番号を出しても、別の番号を探す（GAS 2470行）。"""
    会員を作る("MYM-1001")
    出す = iter([1001, 1001, 2002])
    monkeypatch.setattr(entrance.random, "randrange", lambda *a, **k: next(出す))
    答 = 登録(client)
    assert 答["memberId"] == "MYM-2002"
    assert Member.objects.count() == 2


def test_新規登録_採番できなければ黙って作らず断る(client, monkeypatch):
    会員を作る("MYM-1001")
    monkeypatch.setattr(entrance.random, "randrange", lambda *a, **k: 1001)
    答 = 登録(client)
    assert 答["status"] == "error"
    assert 答["message"] == "会員番号を採番できませんでした。受付にお申し出ください。"
    assert Member.objects.count() == 1


def test_新規登録_登録したパスコードでログインできる(client):
    答 = 登録(client, passcode="123456")
    assert ログイン(client, "鈴木一郎", "123456")["memberId"] == 答["memberId"]


def test_新規登録_氏名と生年月日で受付が作った行を引き受ける(client):
    """受付が先に作った行（パスコード無し）に、お客様があとからパスコードを付ける。"""
    m = 会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05",
                phone="", kana="", memo="受付のメモ", stamp_count=3)
    答 = 登録(client)
    assert 答["status"] == "ok"
    assert 答["memberId"] == "MYM-1001"          # 新しい番号を作らない
    assert Member.objects.count() == 1
    m = 読み直す(m)
    assert check_password("1234", m.passcode_hash)
    # 空いていた欄だけ埋まる
    assert m.kana == "すずきいちろう"
    assert m.phone == "09012345678"
    assert m.address == "平塚市1-2-3"
    # 受付の記録は消えない
    assert m.memo == "受付のメモ"
    assert m.stamp_count == 3


def test_新規登録_引き受けるとき入っている欄は上書きしない(client):
    m = 会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05",
                phone="0463111111", address="寒川町")
    登録(client, phone="09099999999", address="別の住所")
    m = 読み直す(m)
    assert m.phone == "0463111111"
    assert m.address == "寒川町"


def test_新規登録_氏名は空白と全角半角を無視して照らす(client):
    会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05")
    答 = 登録(client, name="鈴木　一郎")
    assert 答["memberId"] == "MYM-1001"


def test_新規登録_生年月日はスラッシュ区切りでも照らす(client):
    会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05")
    答 = 登録(client, birthday="1990/04/05")
    assert 答["memberId"] == "MYM-1001"


def test_新規登録_すでにパスコードがある方は登録ではなくログインへ(client):
    会員を作る("MYM-1001", name="鈴木一郎", passcode="9999", birthday="1990-04-05")
    答 = 登録(client)
    assert 答 == {"status": "error", "message": "すでに登録済みです。パスコードでログインしてください。"}


def test_新規登録_同じ氏名と生年月日が2件あれば受付へ(client):
    会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05")
    会員を作る("MYM-1002", name="鈴木一郎", passcode="", birthday="1990-04-05")
    答 = 登録(client)
    assert 答["message"] == "同じお名前と生年月日の登録が複数あります。受付にお申し出ください。"


def test_新規登録_退会の印がある行は照らさない(client):
    """GAS の readAccountRows_（2224行）は削除済みの行を読まない。新しい番号になる。"""
    会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday="1990-04-05", deleted=True)
    答 = 登録(client)
    assert 答["status"] == "ok"
    assert 答["memberId"] != "MYM-1001"
    assert Member.objects.count() == 2


@pytest.mark.parametrize("項目, 文言", [
    ({"name": ""}, "お名前を入力してください。"),
    ({"birthday": ""}, "生年月日を入力してください。"),
    ({"passcode": "12345"}, "パスコードは数字4桁または6桁で入力してください。"),
    ({"passcode": "abcd"}, "パスコードは数字4桁または6桁で入力してください。"),
    ({"kana": ""}, "フリガナを入力してください。"),
    ({"phone": ""}, "電話番号を入力してください。"),
    ({"address": ""}, "ご住所を入力してください。"),
])
def test_新規登録_足りない項目の文言はGASと同じ(client, 項目, 文言):
    答 = 登録(client, **項目)
    assert 答 == {"status": "error", "message": 文言}
    assert Member.objects.count() == 0


def test_新規登録_電話番号の消えた先頭の0を戻して保存する(client):
    答 = 登録(client, phone="9012345678")
    assert Member.objects.get(pk=答["memberId"]).phone == "09012345678"


@pytest.mark.xfail(reason="GAS 2378行 needsKanjiName_: ひらがなだけのお名前は入口で止める"
                          "（あとから漢字で登録し直されて二重になるため）。サーバーは通してしまう", strict=True)
def test_新規登録_ひらがなだけのお名前は漢字をお願いする(client):
    答 = 登録(client, name="すずきいちろう")
    assert 答["status"] == "error"
    assert 答["message"].startswith("お名前は漢字でご入力ください。")
    assert Member.objects.count() == 0


@pytest.mark.xfail(reason="GAS 2417行: 同じお名前で生年月日が空の行があれば、本人か確かめようがないので"
                          "受付へ案内する。サーバーは新しい行を足してしまい、同じ方が2行に分かれる", strict=True)
def test_新規登録_同名で生年月日が無い行があれば受付へ(client):
    会員を作る("MYM-1001", name="鈴木一郎", passcode="", birthday=None)
    答 = 登録(client)
    assert 答["status"] == "error"
    assert "生年月日が記録されていないため確認できません" in 答["message"]
    assert Member.objects.count() == 1


@pytest.mark.xfail(reason="GAS 2432行: フリガナと生年月日が同じ行があれば、お名前の書き方だけが違う"
                          "同じ方とみなして受付へ案内する。サーバーは新しい行を足してしまう", strict=True)
def test_新規登録_フリガナと生年月日が同じ行があれば受付へ(client):
    会員を作る("MYM-1001", name="こばやしみか", kana="こばやしみか", passcode="1234",
            birthday="1990-04-05")
    答 = 登録(client, name="小林美香", kana="こばやしみか", birthday="1990-04-05")
    assert 答["status"] == "error"
    assert "フリガナと生年月日が同じご登録がすでにあります" in 答["message"]
    assert Member.objects.count() == 1


@pytest.mark.xfail(reason="GAS 2478行は normalizeStoredName_ / normalizeStoredKana_ で空白を取り"
                          "カタカナをひらがなに寄せて保存する。サーバーは受けた文字のまま保存する"
                          "（入口の画面が空白を取って送るので、いまは表に出にくい）", strict=True)
def test_新規登録_お名前の空白を取りフリガナをひらがなに寄せて保存する(client):
    答 = 登録(client, name="鈴木 一郎", kana="スズキ　イチロウ")
    m = Member.objects.get(pk=答["memberId"])
    assert m.name == "鈴木一郎"
    assert m.kana == "すずきいちろう"


@pytest.mark.xfail(reason="GAS 2487行 setUserRegistrationSource_(row, '新規登録', 'ランチャー')。"
                          "サーバーの 新規登録 は登録経路を入れないので、分析の登録経路に出ない", strict=True)
def test_新規登録_登録経路を新規登録_ランチャーで残す(client):
    答 = 登録(client)
    m = Member.objects.get(pk=答["memberId"])
    assert m.registration_source == "新規登録"
    assert m.registration_source_detail == "ランチャー"
    assert m.registration_source_updated_at is not None


@pytest.mark.xfail(reason="GAS 2497行は buildAccountSession_ を返す（token / memberToken / role …）。"
                          "サーバーは memberId / name / kana / claimed だけ。入口は token が無いと"
                          "「ログインできませんでした」にする（start/index.html 677行）", strict=True)
def test_新規登録_入口が要る項目をそろえて返す(client):
    答 = 登録(client)
    assert 答["role"] == "member"
    assert 答["token"] and 答["memberToken"]
    assert 答["bijiris"] is False


# ═════════════════════════════════════════════════════════
# registerBijirisUse
#
# 入口（start/index.html 925行）は **name / kana / passcode** を送ってくる。
# GAS 2594行も、お名前とパスコードでご本人を確かめてから印を付ける。
# ═════════════════════════════════════════════════════════

def ビジリス登録(client, **項目):
    中 = {"type": "registerBijirisUse", "name": "佐藤花子", "kana": "さとうはなこ", "passcode": "1234"}
    中.update(項目)
    return 書く(client, 中).json()


@pytest.mark.xfail(reason="entrance.ビジリス登録 は memberId で探す（誰も送ってこない）。"
                          "入口は name / kana / passcode を送る（start/index.html 925行・GAS 2597行）。"
                          "member を渡すと、ビジリスの利用登録が「会員IDが指定されていません」で全員止まる", strict=True)
def test_ビジリス登録_お名前とパスコードで本人を確かめて印を付ける(client):
    m = 会員を作る(name="佐藤花子", passcode="1234")
    答 = ビジリス登録(client)
    assert 答["status"] == "ok"
    assert 答["memberId"] == "MYM-1001"
    assert 答["bijiris"] is True
    assert 読み直す(m).bijiris_registered is True


@pytest.mark.xfail(reason="同上。GAS 2619行: フリガナが未登録の会員は、この機会にひらがなで埋める", strict=True)
def test_ビジリス登録_フリガナが空なら埋める(client):
    m = 会員を作る(name="佐藤花子", kana="", passcode="1234")
    ビジリス登録(client, kana="サトウハナコ")
    assert 読み直す(m).kana == "さとうはなこ"


@pytest.mark.xfail(reason="同上。GAS 2612行の文言", strict=True)
def test_ビジリス登録_パスコードが違えば印を付けない(client):
    m = 会員を作る(name="佐藤花子", passcode="1234")
    答 = ビジリス登録(client, passcode="0000")
    assert 答 == {"status": "error", "message": "お名前またはパスコードが違います。"}
    assert 読み直す(m).bijiris_registered is False


@pytest.mark.xfail(reason="同上。GAS 2601行の文言", strict=True)
def test_ビジリス登録_足りなければ入力を促す(client):
    答 = ビジリス登録(client, kana="")
    assert 答["message"] == "お名前・フリガナ・パスコードを入力してください。"


def test_ビジリス登録_退会の印がある方には付けない(client):
    """いまの形（memberId）でも、退会者に印を付けないことは守られている。"""
    m = 会員を作る(name="佐藤花子", passcode="1234", deleted=True)
    答 = 書く(client, {"type": "registerBijirisUse", "memberId": "MYM-1001"}).json()
    assert 答["status"] == "error"
    assert 読み直す(m).bijiris_registered is False
