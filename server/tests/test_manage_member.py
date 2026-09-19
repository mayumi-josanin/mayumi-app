"""会員管理。**中身は旧管理アプリ（#page-users 会員一覧）と同じ**
（集計カード5つ・検索と市区町村と年齢とPush状態の絞り込み・並び替え6種・表の列・編集・ボタン3つ・削除）。

会員の書き込みは 9/19 の切り替え（member_gate）まで断る。書く試験は先に切り替える。
"""

from datetime import date, timedelta

import pytest
from django.contrib.auth.hashers import check_password
from django.utils import timezone

from apps.manage import member_gate, views_member
from apps.members.models import Member

pytestmark = pytest.mark.django_db


def _member(mid, name, **extra):
    d = {"kana": "", "phone": "", "address": "", "created_at": timezone.now()}
    d.update(extra)
    return Member.objects.create(member_id=mid, name=name, **d)


@pytest.fixture
def サーバーが正():
    member_gate.切り替える("server")


def test_まゆみ以外は入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/members/").status_code == 403


def test_何も無いときの文言と集計カード(as_owner):
    page = as_owner.get("/manage/members/").content.decode()
    for label in ["会員一覧", "🔄 更新", "会員総数", "Push許可あり", "未対応注文あり", "Push許可なし", "現在表示中",
                  "会員ID・氏名・フリガナ・電話番号・住所・メモで検索", "市区町村: 全て", "年齢 下限", "年齢 上限",
                  "並び替え: 登録が新しい順", "並び替え: 登録が古い順", "並び替え: 年齢が高い順", "並び替え: 年齢が低い順",
                  "並び替え: 市区町村順", "並び替え: 50音順（フリガナ）", "Push状態: 全て", "Push許可あり", "Push許可なし",
                  "条件をクリア", "表示する列を選ぶ", "会員データはありません", "登録されている会員はいません"]:
        assert label in page, label
    # 表の列（旧アプリと同じ12列）
    for th in ["アイコン", "ID", "登録/更新日時", "氏名", "電話番号", "生年月日", "年齢", "住所", "Push", "アンケート", "メモ", "操作"]:
        assert f">{th}</th>" in page, th
    # 左メニューは「カレンダー」の次に「会員管理」
    assert 'class="active">会員一覧</a>' in page and page.index("会員一覧</a>") < page.index("スタンプ・特典</a>")  # 上部タブ
    # 切り替え前は「見るだけ」と出す
    assert "まだスプレッドシートが正です" in page


def test_一覧は旧アプリと同じ計算で出す(as_owner):
    きょう = timezone.localdate()
    _member("MYM-0001", "山田花子", kana="やまだはなこ", phone="09011112222", address="神奈川県厚木市中町1-1",
            birthday=きょう.replace(year=きょう.year - 30), push_subscription="sub-abc",
            survey_answered_at=timezone.now(), memo="覚え書き", device_sessions=[{"deviceId": "d1", "lastSeenAt": timezone.now().isoformat()}])
    _member("MYM-0002", "佐藤桃子", kana="さとうももこ", address="平塚市八重咲町2-2", created_at=timezone.now() - timedelta(days=1),
            reward_history=[{"rewardName": "A賞", "used": False}, {"rewardName": "B賞", "used": True}])
    _member("MYM-0003", "退会さん", deleted=True)
    page = as_owner.get("/manage/members/").content.decode()
    assert "退会さん" not in page
    assert "2 / 2 件を表示" in page
    assert "30歳" in page and "厚木市" in page and "平塚市" in page
    assert "🔔" in page and "最近オンライン" in page and "未同期" in page
    assert "回答 " in page and "スタンプ 未付与" in page and "未回答" in page
    assert "注文 0件 / 受付中 0件 / 端末 1台" in page and "最終注文: ---" in page
    assert "厚木市（1）" in page and "平塚市（1）" in page
    # 登録が新しい順が既定（山田が上）
    assert page.index("MYM-0001") < page.index("MYM-0002")
    page = as_owner.get("/manage/members/?sort=registered_asc").content.decode()
    assert page.index("MYM-0002") < page.index("MYM-0001")
    # 50音順: フリガナが「さ」の佐藤が先
    page = as_owner.get("/manage/members/?sort=kana_asc").content.decode()
    assert page.index("MYM-0002") < page.index("MYM-0001")
    # 検索（長音・空白・大文字小文字を無視）
    page = as_owner.get("/manage/members/?q=0901111").content.decode()
    assert "MYM-0001" in page and "MYM-0002" not in page and "1 / 2 件を表示" in page
    # 市区町村・年齢・Push状態で絞る
    assert "MYM-0002" not in as_owner.get("/manage/members/?city=厚木市").content.decode()
    page = as_owner.get("/manage/members/?age_min=20&age_max=40").content.decode()
    assert "MYM-0001" in page and "MYM-0002" not in page  # 生年月日が無い方は年齢で絞る間は除外
    page = as_owner.get("/manage/members/?push=disabled").content.decode()
    assert "MYM-0002" in page and "MYM-0001" not in page
    page = as_owner.get("/manage/members/?q=いない").content.decode()
    assert "条件に一致する会員はいません" in page


def test_表示する列を選べる(as_owner):
    _member("MYM-0001", "山田花子", kana="やまだはなこ", line_user_id="U-abc", role="管理者", stamp_count=3)
    page = as_owner.get("/manage/members/").content.decode()
    assert ">フリガナ</th>" not in page and "U-abc" not in page
    page = as_owner.get("/manage/members/?cols_given=1&col=memberId&col=kana&col=lineUserId&col=role&col=stamp").content.decode()
    assert ">フリガナ</th>" in page and "やまだはなこ" in page and "U-abc" in page and "管理者" in page and "3個" in page
    assert ">電話番号</th>" not in page and ">操作</th>" in page


def test_年齢と市区町村の計算は旧アプリと同じ():
    きょう = date(2026, 9, 15)
    assert views_member._年齢("1988/4/22", きょう) == 38
    assert views_member._年齢("1988-09-16", きょう) == 37  # 誕生日の前日
    assert views_member._年齢("1988年9月15日", きょう) == 38
    assert views_member._年齢("1988/2/30", きょう) is None
    assert views_member._年齢("", きょう) is None
    assert views_member._市区町村("神奈川県厚木市中町1-1") == "厚木市"
    assert views_member._市区町村("厚木市 中町") == "厚木市"
    assert views_member._市区町村("奈川県高座郡寒川町一之宮") == "寒川町"
    assert views_member._市区町村("東京都世田谷区") == "世田谷区"
    assert views_member._市区町村("") == ""


def test_切り替え前は書き込みを断る(as_owner):
    m = _member("MYM-0001", "山田花子", phone="09011112222")
    r = as_owner.post("/manage/members/MYM-0001/", {"name": "山田花", "phone": "1"}, follow=True)
    assert "まだスプレッドシートが正です" in r.content.decode()
    m.refresh_from_db()
    assert m.name == "山田花子" and m.phone == "09011112222"
    for path in ["passcode/", "transfer-code/", "stop-push/", "delete/"]:
        r = as_owner.post(f"/manage/members/MYM-0001/{path}", {"passcode": "1234"}, follow=True)
        assert "まだスプレッドシートが正です" in r.content.decode(), path
    m.refresh_from_db()
    assert m.deleted is False and m.passcode_hash == "" and m.transfer_code == ""


def test_編集画面の項目(as_owner):
    _member("MYM-0001", "山田花子", kana="やまだはなこ", phone="09011112222", merged_into_id="", transfer_code="12345678",
            transfer_code_issued_at=timezone.now(), password_hash="secret-hash", password_salt="secret-salt",
            passcode_hash="pbkdf2$secret", push_subscription="sub-abc")
    page = as_owner.get("/manage/members/MYM-0001/").content.decode()
    for label in ["氏名", "フリガナ", "電話番号", "生年月日", "住所", "メモ", "現在スタンプ数", "スタンプカード番号",
                  "最終スタンプ取得日", "最終スタンプ取得日時", "権限", "ビジリス", "LINEユーザーID", "ご状況", "登録経路",
                  "🔑 パスコードを設定", "📱 引き継ぎコードを発行", "🔕 通知を止める",
                  "会員番号", "登録日時", "最終オンライン", "統合先会員ID", "引き継ぎコード", "12345678", "通知の届け先", "設定あり",
                  "この会員を削除する"]:
        assert label in page, label
    # 出さないもの: パスワードハッシュ・ソルト・パスコードのハッシュ・届け先そのもの
    for secret in ["secret-hash", "secret-salt", "pbkdf2$secret", "sub-abc"]:
        assert secret not in page, secret
    assert as_owner.get("/manage/members/MYM-9999/").status_code == 404


def test_保存は会員を書き換えるを通す(as_owner, サーバーが正):
    _member("MYM-0001", "山田花子")
    _member("MYM-0002", "佐藤桃子", line_user_id="U-taken")
    r = as_owner.post("/manage/members/MYM-0001/", {
        "name": "山田 花子", "kana": "ヤマダ ハナコ", "phone": "9011112222", "birthday": "1988-04-22", "address": "厚木市",
        "memo": "メモ", "status": "産後", "role": "管理者", "bijiris": "登録済み", "stampCount": "4", "stampCardNum": "2",
        "lastStampAt": "2026-09-01T10:30", "registrationSource": "紹介", "registrationSourceDetail": "友人", "lineUserId": "U-new",
    })
    assert r.status_code == 302
    m = Member.objects.get(pk="MYM-0001")
    assert m.name == "山田花子" and m.kana == "やまだはなこ"
    assert m.phone == "09011112222"  # 先頭の0を戻す
    assert m.birthday == date(1988, 4, 22) and m.address == "厚木市" and m.memo == "メモ" and m.status == "産後"
    assert m.role == "管理者" and m.bijiris_registered is True and m.stamp_count == 4 and m.stamp_card_number == 2
    assert timezone.localtime(m.last_stamp_at).strftime("%Y-%m-%d %H:%M") == "2026-09-01 10:30"
    assert m.registration_source == "紹介" and m.registration_source_detail == "友人" and m.line_user_id == "U-new"
    # 開き直しても残る
    page = as_owner.get("/manage/members/MYM-0001/").content.decode()
    assert 'value="09011112222"' in page and 'value="やまだはなこ"' in page and 'value="2026-09-01T10:30"' in page
    # LINEユーザーIDの重複は断る
    r = as_owner.post("/manage/members/MYM-0001/", {"name": "山田花子", "lineUserId": "U-taken"})
    assert r.status_code == 200 and "ほかの会員（MYM-0002）に付いています" in r.content.decode()
    assert Member.objects.get(pk="MYM-0001").line_user_id == "U-new"
    # 氏名は空にできない
    r = as_owner.post("/manage/members/MYM-0001/", {"name": " "})
    assert "氏名を入力してください" in r.content.decode()


def test_パスコードは設定できるが値は出ない(as_owner, サーバーが正):
    _member("MYM-0001", "山田花子", password_hash="old", password_salt="salt")
    r = as_owner.post("/manage/members/MYM-0001/passcode/", {"passcode": "123"}, follow=True)
    assert "数字4桁または6桁" in r.content.decode()
    r = as_owner.post("/manage/members/MYM-0001/passcode/", {"passcode": "4321"}, follow=True)
    page = r.content.decode()
    assert "パスコードを設定しました" in page and "4321" not in page
    m = Member.objects.get(pk="MYM-0001")
    assert check_password("4321", m.passcode_hash) and m.password_hash == "" and m.password_salt == ""
    assert "4321" not in as_owner.get("/manage/members/").content.decode()


def test_引き継ぎコードは発行の結果を画面に出す(as_owner, サーバーが正):
    _member("MYM-0001", "山田花子")
    r = as_owner.post("/manage/members/MYM-0001/transfer-code/", follow=True)
    m = Member.objects.get(pk="MYM-0001")
    page = r.content.decode()
    assert len(m.transfer_code) == 8 and m.transfer_code.isdigit()
    assert f"引き継ぎコード: {m.transfer_code}" in page and "有効期限: " in page and "お客様にお伝えください" in page
    assert timezone.localtime(m.transfer_code_issued_at + timedelta(hours=168)).strftime("%Y年%-m月%-d日 %H:%M") in page


def test_通知を止める(as_owner, サーバーが正):
    _member("MYM-0001", "山田花子", push_subscription="sub-abc", push_enabled=True)
    r = as_owner.post("/manage/members/MYM-0001/stop-push/", follow=True)
    assert "通知を止めました" in r.content.decode()
    m = Member.objects.get(pk="MYM-0001")
    assert m.push_subscription == "" and m.push_enabled is False


def test_削除は印だけで行は残る(as_owner, サーバーが正):
    _member("MYM-0001", "山田花子")
    r = as_owner.post("/manage/members/MYM-0001/delete/", follow=True)
    assert "会員を削除しました" in r.content.decode()
    m = Member.objects.get(pk="MYM-0001")
    assert m.deleted is True and m.deleted_at is not None
    assert "MYM-0001" not in as_owner.get("/manage/members/").content.decode()
    assert as_owner.get("/manage/members/MYM-0001/").status_code == 404  # 消した会員は編集しない


def test_一覧の1行の縦幅はどの行も同じ(as_owner):
    """行の中身が多い会員も少ない会員も、同じ高さで並ぶこと（院長の希望 2026-09-19）。

    氏名は3行・住所とメモとアンケートは2行あり、**中身の行数で高さが変わっていた。**
    """
    _member("MYM-0001", "山田花子", address="神奈川県厚木市中町1-1-1 コーポ山田303号室", memo="長い覚え書き" * 10)
    _member("MYM-0002", "佐藤桃子")
    page = as_owner.get("/manage/members/").content.decode()
    assert 'class="member-table"' in page
    # はみ出す文字は「…」で省く（1行に収める印）。全文はホバーで出す
    assert "member-line" in page
    assert 'title="神奈川県厚木市中町1-1-1 コーポ山田303号室"' in page
