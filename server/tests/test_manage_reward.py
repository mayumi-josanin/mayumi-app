"""スタンプ・特典管理。**中身は旧管理アプリ（#page-stamp-rewards）と同じ**
（月別ガチャ特典設定・集計カード4つ・検索と2つの絞り込み・表10列・詳細・編集）。"""

from datetime import datetime

import pytest
from django.utils import timezone

from apps.manage import member_gate
from apps.members.models import Member
from apps.records.models import AppSetting

pytestmark = pytest.mark.django_db


def 日本時間(*a):
    return timezone.make_aware(datetime(*a))


def _会員(member_id="MYM-1001", name="佐藤花子", **項目):
    項目.setdefault("device_sessions", [])
    return Member.objects.create(member_id=member_id, name=name, **項目)


def _特典(**上書き):
    r = {"id": "r1", "cardNum": 1, "rewardName": "A賞 よもぎ茶", "rewardNote": "受付でお受け取りください",
         "earnedDate": "2026-04-01T10:00:00+09:00", "expiryDate": "2099-05-01T23:59:59+09:00", "used": False, "usedAt": ""}
    r.update(上書き)
    return r


# ═════════════════════════════════════════════════════════
# 一覧
# ═════════════════════════════════════════════════════════

def test_会員がいないときの文言と集計カード(as_owner):
    page = as_owner.get("/manage/rewards/").content.decode()
    for label in ["スタンプ・特典管理", "🎁 月別ガチャ特典設定", "🎟 会員別スタンプ・特典状況",
                  "会員総数", "達成済みカード", "未受取特典数", "現在表示中",
                  "会員ID・氏名・電話番号・住所で検索", "条件をクリア", "登録されている会員はいません", "会員データはありません",
                  "最終スタンプ取得日時が新しい会員から上に表示します"]:
        assert label in page, label
    # 旧アプリの「特典データから会員を再作成」は GAS に受け口が無いので作らない
    assert "会員を再作成" not in page
    # 左メニュー
    assert "スタンプ・特典</a>" in page


def test_表の列と受け取り状況と状態(as_owner):
    _会員(stamp_count=10, stamp_card_number=2, last_stamp_at=日本時間(2026, 4, 3, 9, 0),
        reward_history=[_特典(), _特典(id="r2", used=True, usedAt="2026-04-05T12:00:00+09:00"),
                        _特典(id="r3", earnedDate="2026-03-01T10:00:00+09:00", expiryDate="2026-03-31T23:59:59+09:00")])
    page = as_owner.get("/manage/rewards/").content.decode()
    for th in ["会員ID", "氏名 / 電話", "カード", "スタンプ", "最終スタンプ", "受け取り状況", "有効期限", "最新特典", "状態", "操作"]:
        assert f"<th>{th}</th>" in page, th
    assert "2枚目" in page and "10 / 10" in page and "2026/04/03" in page
    assert "未受取 1" in page and "受取済み 1" in page and "期限切れ 1" in page and "総数 3" in page
    assert "2099/05/01" in page  # 未受取で期限内のいちばん近い有効期限
    assert "2026/04/01" in page  # 最新特典（獲得日）
    assert "達成済み" in page and "✓ 単独" in page
    assert "1 / 1 件を表示" in page
    assert "/manage/rewards/MYM-1001/" in page


def test_絞り込みと検索(as_owner):
    _会員("MYM-1001", "達成さん", stamp_count=10, phone="09011112222", reward_history=[_特典()])
    _会員("MYM-1002", "途中さん", stamp_count=3, address="平塚市", reward_history=[_特典(used=True)])
    _会員("MYM-1003", "未取得さん", stamp_count=0)
    page = as_owner.get("/manage/rewards/?stamp=completed").content.decode()
    assert "達成さん" in page and "途中さん" not in page and "未取得さん" not in page
    page = as_owner.get("/manage/rewards/?stamp=collecting").content.decode()
    assert "途中さん" in page and "達成さん" not in page
    page = as_owner.get("/manage/rewards/?stamp=empty").content.decode()
    assert "未取得さん" in page and "途中さん" not in page
    page = as_owner.get("/manage/rewards/?ticket=unused").content.decode()
    assert "達成さん" in page and "途中さん" not in page
    page = as_owner.get("/manage/rewards/?ticket=earned").content.decode()
    assert "達成さん" in page and "途中さん" in page and "未取得さん" not in page
    page = as_owner.get("/manage/rewards/?ticket=none").content.decode()
    assert "未取得さん" in page and "達成さん" not in page
    # 電話番号・住所でも探せる（旧アプリと同じく、空白は無視・ハイフンはそのまま）
    page = as_owner.get("/manage/rewards/?q=0901111").content.decode()
    assert "達成さん" in page and "途中さん" not in page
    page = as_owner.get("/manage/rewards/?q=平塚").content.decode()
    assert "途中さん" in page and "達成さん" not in page
    page = as_owner.get("/manage/rewards/?q=いない").content.decode()
    assert "条件に一致する会員はいません" in page and "0 / 3 件を表示" in page
    # 集計カードは絞り込みに関わらず全員（現在表示中だけ絞った数）
    assert "会員総数" in page


def test_最終スタンプ取得日時が新しい会員から上(as_owner):
    _会員("MYM-1001", "古い", last_stamp_at=日本時間(2026, 1, 1))
    _会員("MYM-1002", "新しい", last_stamp_at=日本時間(2026, 5, 1))
    _会員("MYM-1003", "動き無し")
    page = as_owner.get("/manage/rewards/").content.decode()
    assert page.index("新しい") < page.index("古い") < page.index("動き無し")


def test_重複候補の印(as_owner):
    _会員("MYM-1001", "佐藤花子", phone="09011112222")
    _会員("MYM-1002", "佐藤 花子", phone="090-1111-2222")
    _会員("MYM-1003", "鈴木一郎")
    page = as_owner.get("/manage/rewards/").content.decode()
    assert page.count("⚠ 重複あり") == 2 and page.count("✓ 単独") == 1
    assert "電話番号が一致 / 氏名が一致" in page


# ═════════════════════════════════════════════════════════
# 月別ガチャ特典設定
# ═════════════════════════════════════════════════════════

def _ガチャ(i, month, a=5, b=15, c=30, d=50, **中身):
    出 = {f"month_{i}": month}
    for k, p in zip("ABCD", (a, b, c, d)):
        出[f"probability_{i}_{k}"] = str(p)
        出[f"content_{i}_{k}"] = 中身.get(k, "")
        出[f"note_{i}_{k}"] = ""
    return 出


def test_設定が無いときは今月の既定の行が出る(as_owner):
    page = as_owner.get("/manage/rewards/").content.decode()
    今月 = timezone.localdate().strftime("%Y-%m")
    assert f'name="month_0" class="form-control gacha-month" value="{今月}"' in page
    assert 'name="probability_0_A" class="form-control gacha-prob" value="5"' in page
    assert "合計 100%" in page and "➕ 月を追加" in page and "💾 保存" in page
    assert f"現在月は {今月[:4]}年{int(今月[5:])}月 です。抽選では" in page


def test_保存すると月順に並び確率は小さい順にAからDへ(as_owner):
    r = as_owner.post("/manage/rewards/gacha/", {
        "row": ["0", "1"],
        **_ガチャ(0, "2026-05", a=50, b=30, c=15, d=5, A="お茶"),
        **_ガチャ(1, "2026-04", A="石けん"),
    })
    assert r.status_code == 302
    値 = AppSetting.objects.get(pk="REWARD_GACHA_CONFIG").value
    assert [e["month"] for e in 値["monthlyPrizes"]] == ["2026-04", "2026-05"]
    五月 = 値["monthlyPrizes"][1]["prizes"]
    # GAS の normalizeRewardGachaPrizes_ と同じ: A賞がいちばん出にくい並びに直す
    assert [五月[k]["probability"] for k in "ABCD"] == [5, 15, 30, 50] and 五月["A"]["content"] == "お茶"
    page = as_owner.get("/manage/rewards/").content.decode()
    assert "月別ガチャ特典設定を保存しました" in page and "石けん" in page and 'value="2026-04"' in page
    # 会員の表ではないので、切り替え前でも保存できる
    assert not member_gate.会員はサーバーが正()


def test_確率の合計が100でなければ保存しない(as_owner):
    r = as_owner.post("/manage/rewards/gacha/", {"row": ["0"], **_ガチャ(0, "2026-04", a=10)})
    assert r.status_code == 200
    page = r.content.decode()
    assert "2026年4月 の出現確率合計を 100% にしてください。" in page
    assert 'name="probability_0_A" class="form-control gacha-prob" value="10"' in page  # 入力は残す
    assert not AppSetting.objects.filter(pk="REWARD_GACHA_CONFIG").exists()


def test_同じ月が2行あれば保存しない(as_owner):
    r = as_owner.post("/manage/rewards/gacha/", {"row": ["0", "1"], **_ガチャ(0, "2026-04"), **_ガチャ(1, "2026-04")})
    assert "同じ月が重複しています。月ごとに1行だけ設定してください。" in r.content.decode()
    r = as_owner.post("/manage/rewards/gacha/", {"row": ["0"], **_ガチャ(0, "")})
    assert "対象月を選択してください。" in r.content.decode()
    r = as_owner.post("/manage/rewards/gacha/", {})
    assert "設定する月がありません。" in r.content.decode()


# ═════════════════════════════════════════════════════════
# 詳細・編集
# ═════════════════════════════════════════════════════════

def test_詳細と編集画面の項目(as_owner):
    _会員(stamp_count=4, stamp_card_number=2, last_stamp_at=日本時間(2026, 4, 3, 9, 30), stamp_achieved_at=日本時間(2026, 3, 20, 15, 0),
        reward_history=[_特典(cardNum=1, used=True, usedAt="2026-04-05T12:00:00+09:00")])
    page = as_owner.get("/manage/rewards/MYM-1001/").content.decode()
    for label in ["🔎 会員のスタンプ・特典詳細", "現在カード", "現在スタンプ", "未受取特典", "受取済み特典", "カード別スタンプ状況",
                  "1枚目のスタンプカード", "2枚目のスタンプカード", "記録 / 特典内容", "受け取り状況", "有効期限", "受取日時",
                  "🎟️ スタンプ・特典状況の編集", "現在のスタンプ数", "スタンプカード番号", "最終スタンプ取得日", "達成日時",
                  "獲得済み特典", "➕ 特典を追加", "受取済みにする", "説明書き（注意書き）"]:
        assert label in page, label
    assert "2枚目</div>" in page and "4 / 10" in page
    assert 'name="lastStampDate" type="date" class="form-control" value="2026-04-03"' in page
    assert 'name="stampAchievedDate" type="datetime-local" class="form-control" value="2026-03-20T15:00"' in page
    assert "1枚目カードの特典履歴 1" in page and "受取済み" in page and 'name="r_used_0" checked' in page
    assert "このカードに紐づく特典履歴はありません" in page  # 2枚目（現在・収集中）には特典が無い
    assert as_owner.get("/manage/rewards/MYM-9999/").status_code == 404


def test_達成済みで特典未発行の文言(as_owner):
    _会員(stamp_count=10, stamp_card_number=1)
    page = as_owner.get("/manage/rewards/MYM-1001/").content.decode()
    assert "達成済み / 特典未発行" in page and "達成済みですが、まだ特典は発行されていません" in page


def _編集の入力():
    return {
        "stampCount": "7", "stampCardNum": "2", "lastStampDate": "2026-04-10", "stampAchievedDate": "2026-04-01T10:30",
        "reward": ["0", "1"],
        "r_id_0": "r1", "r_card_0": "1", "r_earned_0": "2026-04-01", "r_expiry_0": "", "r_name_0": "A賞 よもぎ茶", "r_note_0": "受付で",
        "r_used_at_0": "2026-04-05T12:00",
        "r_id_1": "", "r_card_1": "2", "r_earned_1": "2026-04-08", "r_expiry_1": "2026-05-31", "r_name_1": "", "r_note_1": "",
        "r_used_at_1": "", "r_used_1": "on",
    }


def test_切り替え前は会員ごとの編集を断る(as_owner):
    m = _会員(stamp_count=2)
    r = as_owner.post("/manage/rewards/MYM-1001/", _編集の入力())
    assert r.status_code == 302
    page = as_owner.get("/manage/rewards/").content.decode()
    assert member_gate.断る文() in page
    m.refresh_from_db()
    assert m.stamp_count == 2 and m.reward_admin_set_at is None


def test_切り替え後は編集が保存される(as_owner):
    member_gate.切り替える("server")
    m = _会員(stamp_count=2, last_stamp_at=日本時間(2026, 4, 10, 9, 15))
    r = as_owner.post("/manage/rewards/MYM-1001/", _編集の入力())
    assert r.status_code == 302
    m.refresh_from_db()
    assert m.stamp_count == 7 and m.stamp_card_number == 2
    # 最終スタンプ取得日は日付だけ送る。日が同じなら時刻は残す
    assert timezone.localtime(m.last_stamp_at) == 日本時間(2026, 4, 10, 9, 15)
    assert timezone.localtime(m.stamp_achieved_at) == 日本時間(2026, 4, 1, 10, 30)
    assert m.reward_admin_set_at is not None
    # 特典の一覧は GAS の sanitizeRewardList_ と同じ姿（獲得日の新しい順・期限の既定は獲得日の1か月後・受取日時があれば受取済み）
    新, 旧 = m.reward_history
    assert 新["cardNum"] == 2 and 新["rewardName"] == "スタンプ達成特典" and 新["used"] is True and 新["usedAt"] == ""
    assert 新["earnedDate"] == "2026-04-08T00:00:00+09:00" and 新["expiryDate"] == "2026-05-31T23:59:59+09:00"
    assert 新["id"].startswith("reward-")
    assert 旧["id"] == "r1" and 旧["used"] is True and 旧["usedAt"] == "2026-04-05T12:00:00+09:00" and 旧["rewardNote"] == "受付で"
    assert 旧["expiryDate"] == "2026-05-01T00:00:00+09:00"
    assert "スタンプ・特典状況を更新しました" in as_owner.get("/manage/rewards/").content.decode()


def test_最終スタンプ取得日を変えると日付が動き空にすると消える(as_owner):
    member_gate.切り替える("server")
    m = _会員(stamp_count=2, last_stamp_at=日本時間(2026, 4, 10, 9, 15))
    as_owner.post("/manage/rewards/MYM-1001/", {**_編集の入力(), "reward": [], "lastStampDate": "2026-04-12"})
    m.refresh_from_db()
    assert timezone.localtime(m.last_stamp_at) == 日本時間(2026, 4, 12, 0, 0) and m.reward_history == []
    as_owner.post("/manage/rewards/MYM-1001/", {**_編集の入力(), "reward": [], "lastStampDate": ""})
    m.refresh_from_db()
    assert m.last_stamp_at is None


# ═════════════════════════════════════════════════════════
# 一覧の欄をその場で直す（院長の依頼 2026-09-21）
# ═════════════════════════════════════════════════════════

def test_一覧の欄がその場で直せる形で出る(as_owner):
    member_gate.切り替える("server")
    _会員(stamp_count=3, stamp_card_number=2, last_stamp_at=日本時間(2026, 4, 3, 9, 0))
    page = as_owner.get("/manage/rewards/").content.decode()
    # 行ごとの保存のあて先（表の中に form は置けないので外に出す）
    assert 'action="/manage/rewards/MYM-1001/row-save/"' in page
    assert 'id="rrow-MYM-1001"' in page
    # 直せるのはカード・スタンプ・最終スタンプの3つ
    assert 'name="stampCardNum"' in page and 'value="2"' in page
    assert 'name="stampCount"' in page
    assert 'name="lastStampDate"' in page and 'value="2026-04-03"' in page
    assert page.count('class="cell-input"') == 3
    assert "cell-edit" in page and ">保存</button>" in page
    # 氏名は会員管理へ（ここでは直さない）
    assert '/manage/members/MYM-1001/' in page
    # 絞り込みを残して戻れるよう、いまの画面の住所を持たせる
    assert 'name="next" value="/manage/rewards/"' in page


def test_会員IDと受け取り状況は直せない(as_owner):
    member_gate.切り替える("server")
    _会員(stamp_count=3, reward_history=[_特典()])
    page = as_owner.get("/manage/rewards/").content.decode()
    # 会員番号・特典の数・状態を送る入力欄は作らない（計算で出るもの・会員を指す鍵）
    for name in ["memberId", "member_id", "used", "unused", "counts", "expiryDate"]:
        assert f'name="{name}"' not in page, name


def test_行からカードとスタンプと最終スタンプを保存できる(as_owner):
    member_gate.切り替える("server")
    m = _会員(stamp_count=2, stamp_card_number=1, last_stamp_at=日本時間(2026, 4, 10, 9, 15),
           reward_history=[_特典()])
    r = as_owner.post("/manage/rewards/MYM-1001/row-save/",
                      {"stampCardNum": "3", "stampCount": "6", "lastStampDate": "2026-04-12"})
    assert r.status_code == 302
    m.refresh_from_db()
    assert m.stamp_card_number == 3 and m.stamp_count == 6
    assert timezone.localtime(m.last_stamp_at) == 日本時間(2026, 4, 12, 0, 0)
    assert m.reward_admin_set_at is not None
    # 特典の一覧は行からは触らない（直すのは詳細の画面）
    assert m.reward_history == [_特典()]
    assert "スタンプ・特典状況を更新しました" in as_owner.get("/manage/rewards/").content.decode()


def test_行の保存でスタンプの上限を超えた値は丸められる(as_owner):
    member_gate.切り替える("server")
    m = _会員(stamp_count=2)
    as_owner.post("/manage/rewards/MYM-1001/row-save/",
                  {"stampCardNum": "0", "stampCount": "99", "lastStampDate": ""})
    m.refresh_from_db()
    assert m.stamp_count == 10 and m.stamp_card_number == 1 and m.last_stamp_at is None
    as_owner.post("/manage/rewards/MYM-1001/row-save/",
                  {"stampCardNum": "2", "stampCount": "-5", "lastStampDate": ""})
    m.refresh_from_db()
    assert m.stamp_count == 0


def test_切り替え前は行の保存も断る(as_owner):
    m = _会員(stamp_count=2, stamp_card_number=1)
    r = as_owner.post("/manage/rewards/MYM-1001/row-save/",
                      {"stampCardNum": "3", "stampCount": "6", "lastStampDate": "2026-04-12"})
    assert r.status_code == 302
    m.refresh_from_db()
    assert m.stamp_count == 2 and m.stamp_card_number == 1 and m.reward_admin_set_at is None
    page = as_owner.get("/manage/rewards/").content.decode()
    assert member_gate.断る文() in page
    # 直せないときは入力欄も保存ボタンも出さない
    assert 'class="cell-input"' not in page and "row-save/" not in page and ">保存</button>" not in page
    assert "いまは見るだけです" in page


def test_行の保存のあとも絞り込みが残る(as_owner):
    member_gate.切り替える("server")
    _会員(stamp_count=2)
    戻り = "/manage/rewards/?q=%E4%BD%90%E8%97%A4&stamp=collecting&ticket=all"
    r = as_owner.post("/manage/rewards/MYM-1001/row-save/",
                      {"stampCardNum": "1", "stampCount": "4", "lastStampDate": "", "next": 戻り})
    assert r.status_code == 302 and r["Location"] == 戻り
    # よその住所へは飛ばさない
    r = as_owner.post("/manage/rewards/MYM-1001/row-save/",
                      {"stampCardNum": "1", "stampCount": "4", "lastStampDate": "", "next": "https://example.com/"})
    assert r["Location"] == "/manage/rewards/"


def test_行の保存は他の会員を変えない(as_owner):
    member_gate.切り替える("server")
    _会員("MYM-1001", "佐藤花子", stamp_count=2)
    ほか = _会員("MYM-1002", "鈴木一子", stamp_count=5, stamp_card_number=2,
              last_stamp_at=日本時間(2026, 4, 1, 8, 0))
    as_owner.post("/manage/rewards/MYM-1001/row-save/",
                  {"stampCardNum": "3", "stampCount": "9", "lastStampDate": "2026-04-12"})
    ほか.refresh_from_db()
    assert ほか.stamp_count == 5 and ほか.stamp_card_number == 2
    assert timezone.localtime(ほか.last_stamp_at) == 日本時間(2026, 4, 1, 8, 0)
    assert ほか.reward_admin_set_at is None


def test_スタッフは入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/rewards/").status_code == 403
