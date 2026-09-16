"""ビジリス管理「顧客管理」。**中身は旧ビジリス管理アプリ（#page-customers）と同じ**
（検索・回数券・回答状態の絞り込み → 顧客一覧 → 顧客情報／編集／メモ／マイルストーン特典／測定履歴／回答履歴）。"""

from datetime import date, datetime
from urllib.parse import quote

import pytest
from django.utils import timezone

from apps.bijiris import gate, views_customers
from apps.bijiris.models import CustomerProfile, Response, ResponsePhoto
from apps.measurements.models import Measurement
from apps.members.models import Member
from apps.records.models import AppSetting

pytestmark = pytest.mark.django_db


def 日本時間(*a):
    return timezone.make_aware(datetime(*a))


def _url(name, *rest):
    return "/manage/bijiris/customers/" + quote(name) + "/" + ("/".join(rest) + "/" if rest else "")


def _回答(name, at, response_id, plan="", sheet="", round_="", status="new", timing="", session_type="", survey="施術後アンケート", key="survey_bijiris_session"):
    answers = []
    if plan:
        answers += [{"questionId": "q_bijiris_session_ticket_plan", "value": plan},
                    {"questionId": "q_bijiris_session_ticket_sheet", "value": sheet},
                    {"questionId": "q_bijiris_session_ticket_round", "value": round_}]
    if timing:
        answers.append({"questionId": "q_measure_timing", "value": [timing]})
    if session_type:
        answers.append({"questionId": "q_bijiris_session_type", "value": session_type})
    return Response.objects.create(response_id=response_id, customer_name=name, submitted_at=at, status=status,
                                   survey_key=key, survey_title=survey, answers=answers)


def _花子():
    """10回券を1枚使い切り（回答は9回目まで＋計測時「回数券終了時」）、2枚目の3回目まで進んだ方。"""
    p = CustomerProfile.objects.create(
        name="佐藤花子", member_number="MYM-0012", name_kana="サトウハナコ",
        active_ticket_card={"plan": "10回券", "sheetNumber": 2, "round": 3}, active_ticket_card_source="response",
        push_status={"enabled": True, "supported": True, "permission": "granted", "updatedAt": "2026-08-01T01:00:00.000Z"},
        ticket_stamp_adjustment=1, updated_at=日本時間(2026, 9, 1),
        memo_entries=[{"at": "2026-09-01", "memo": "新しい"}, {"at": "2026-08-01", "memo": "古い"}], latest_memo="新しい")
    for i in range(1, 10):
        _回答("佐藤花子", 日本時間(2026, 5, i, 10, 0), f"r1-{i}", "10回券", "1枚目", f"{i}回目")
    _回答("佐藤花子", 日本時間(2026, 6, 1, 10, 0), "m-end", timing="回数券終了時", survey="計測時アンケート", key="survey_measurement")
    for i in range(1, 4):
        _回答("佐藤花子", 日本時間(2026, 7, i, 10, 0), f"r2-{i}", "10回券", "2枚目", f"{i}回目")
    _回答("佐藤花子", 日本時間(2026, 7, 10, 10, 0), "camp1", session_type="キャンペーン")
    _回答("佐藤花子", 日本時間(2026, 7, 20, 10, 0), "trash1", "10回券", "2枚目", "9回目", status="trash")
    Measurement.objects.create(measurement_id="m1", customer_name="佐藤花子", measured_on=date(2026, 5, 1), waist=70, hip=95, thigh_right=50, thigh_left=49.5, whr=0.737)
    Measurement.objects.create(measurement_id="m2", customer_name="佐藤花子", measured_on=date(2026, 7, 1), waist=68.5, hip=94, thigh_right=49, thigh_left=49, whr=0.729, staff_memo="順調")
    return p


def _特典設定(**上書き):
    v = {"milestoneRewardConfig": {"enabled": True, "milestones": [
        {"threshold": 1, "reward": "お茶", "description": "受付で"}, {"threshold": 3, "reward": "施術1回", "description": ""}]}}
    v.update(上書き)
    AppSetting.objects.update_or_create(pk="bijiris_preferences", defaults={"value": v})


# ═════════════════════════════════════════════════════════
# 一覧
# ═════════════════════════════════════════════════════════

def test_顧客がいないときの文言と絞り込みの項目(as_owner):
    page = as_owner.get("/manage/bijiris/customers/").content.decode()
    for s in ["<h1>顧客管理</h1>", "会員番号・氏名・フリガナ", ">6回券<", ">10回券<", "回答あり", "回答なし",
              "顧客を選ぶと、顧客情報とアンケート回答履歴を表示します。", "まだ顧客データはありません。", "顧客を追加"]:
        assert s in page, s
    assert "準備中" not in page and gate.断る文() in page
    # 左メニューのタブ
    assert 'class="active">顧客管理</a>' in page


def test_一覧はプロフィールと回答のお名前を合わせ新しい回答がある方から上(as_owner):
    _花子()
    CustomerProfile.objects.create(name="回答なし子", member_number="MYM-0003")
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    _回答("ゴミ箱だけ", 日本時間(2026, 8, 2, 9, 0), "g1", status="trash")
    page = as_owner.get("/manage/bijiris/customers/").content.decode()
    for th in ["お名前", "会員番号", "回答数", "最新回答", "操作"]:
        assert f"<th>{th}</th>" in page, th
    # 花子はプロフィールの更新日時（9/1）が最新の回答（7/10）より新しいので上（旧アプリの latestAt と同じ）
    assert page.find(">佐藤花子</a>") < page.find(">山田太郎</a>") < page.find(">回答なし子</a>")
    assert "ゴミ箱だけ" not in page
    assert "MYM-0012" in page and "サトウハナコ" in page and "14件" in page and "2026/07/20" not in page and "2026/07/10 10:00" in page
    assert "3 / 3 名を表示" in page
    # 検索はフリガナ・会員番号でも当たる
    assert "佐藤花子" in as_owner.get("/manage/bijiris/customers/?q=サトウ").content.decode()
    html = as_owner.get("/manage/bijiris/customers/?q=mym-0003").content.decode()
    assert "回答なし子" in html and "佐藤花子" not in html.split("顧客を追加")[0]
    # 回数券・回答状態
    html = as_owner.get("/manage/bijiris/customers/?ticket=10回券").content.decode().split("顧客を追加")[0]
    assert "佐藤花子" in html and "山田太郎" not in html
    html = as_owner.get("/manage/bijiris/customers/?state=without").content.decode().split("顧客を追加")[0]
    assert "回答なし子" in html and "佐藤花子" not in html and "1 / 3 名を表示" in html
    html = as_owner.get("/manage/bijiris/customers/?q=いない").content.decode()
    assert "条件に一致する顧客はいません。" in html


def test_スタッフは403(client, staff):
    client.force_login(staff)
    assert client.get("/manage/bijiris/customers/").status_code == 403
    assert client.get(_url("佐藤花子")).status_code == 403


# ═════════════════════════════════════════════════════════
# 数え方（旧アプリの app.js / Code.gs と同じ）
# ═════════════════════════════════════════════════════════

def test_回数券の数え方():
    _花子()
    p = CustomerProfile.objects.get(name="佐藤花子")
    回答 = views_customers._生きている回答("佐藤花子")
    assert len(回答) == 14 and 回答[0].response_id == "camp1"
    # 顧客情報の「回数券 合計」: 1枚目は9回目までしか回答が無いが、計測時「回数券終了時」で満了扱い
    assert views_customers.満了したカードの内訳(回答) == {"total": 1, "by_plan": {"10回券": 1}}
    # マイルストーンのスタンプ（Code.gs）: 最終回の回答が無いので 0。手当て +1 で合計 1
    assert views_customers.満了したカードの数(回答) == 0
    # 現在の回数券はプロフィールのカードが正
    t = views_customers.現在の回数券(p, 回答)
    assert (t["plan"], t["sheet_label"], t["round_label"], t["total"]) == ("10回券", "2枚目", "3回目", 10)
    stamps = views_customers.回数券のスタンプ(t, 回答)
    assert len(stamps) == 10 and [s["active"] for s in stamps][:4] == [True, True, True, False]
    assert stamps[0]["date"] == "26/07/01" and stamps[3]["date"] == ""
    # プロフィールにカードが無ければ最新の回数券回答から。1枚目の最新は9回目だが「回数券終了時」がその後にあるので満了
    p.active_ticket_card = None
    Response.objects.filter(response_id__startswith="r2-").delete()
    回答 = views_customers._生きている回答("佐藤花子")
    t = views_customers.現在の回数券(p, 回答)
    assert (t["sheet_label"], t["round_label"]) == ("1枚目", "10回目")
    assert views_customers.キャンペーン(回答) == {"completed": 0, "current": 1}
    assert views_customers.現在の回数券(None, []) is None


def test_通知の状態と会員番号の整え方():
    f = views_customers.通知の状態
    assert f(None)["label"] == "未設定" and f({"supported": False})["label"] == "未対応"
    assert f({"enabled": True, "supported": True})["label"] == "オン"
    assert f({"enabled": False, "supported": True, "permission": "denied"})["label"] == "拒否"
    assert f({"enabled": False, "supported": True})["label"] == "オフ"
    assert views_customers.会員番号を整える("m12") == "MYM-0012" and views_customers.会員番号を整える("abc") == ""


# ═════════════════════════════════════════════════════════
# 顧客情報（詳細）
# ═════════════════════════════════════════════════════════

def test_顧客情報の中身(as_owner):
    _花子()
    _特典設定()
    Member.objects.create(member_id="MYM-0012", name="佐藤花子")
    page = as_owner.get(_url("佐藤花子")).content.decode()
    # 顧客情報（renderCustomerSummaryCard）
    for s in ["<h2>MYM-0012 / 佐藤花子</h2>", "フリガナ: サトウハナコ", ">オン</span> 顧客アプリで通知受信が有効です。",
              "回答数: 14件 / アンケート種類: 2件", "回数券 合計 1枚：6回券 0枚、10回券 1枚",
              "キャンペーン: 完了 0回 / 進行中 1回", "最新回答: 施術後アンケート / 2026/07/10 10:00",
              "最新測定: 2026/07/01 / WHR 0.729", "<strong>2枚目</strong><span>3 / 10</span>"]:
        assert s in page, s
    # お名前が同じ会員がいても、自動では結ばない
    assert "まゆみの会員 MYM-0012" not in page and "（結ばない）" in page and ">MYM-0012 佐藤花子<" in page
    # 顧客情報を編集
    for s in ["パスコードを再設定できるようにする", "パスコード未設定", 'value="MYM-0012"', 'placeholder="自動採番"',
              '<option value="10回券" selected>', '<option value="2枚目" selected>', '<option value="3回目" selected>',
              "顧客情報を保存", "顧客削除", "まゆみの会員と結ぶ"]:
        assert s in page, s
    # 顧客メモ（新しい順）
    assert "メモ履歴 2件" in page and page.find("新しい") < page.find("古い") and "2026/09/01" in page
    # マイルストーン特典
    for s in ["スタンプ: <b>1個</b>（施術後アンケートから 0個 ＋ 手当て +1個）", "道のり（1 / 3）",
              "<td class=\"nowrap\">1個</td>", "お茶", "受付で", "施術1回", "未受取 1件", "未達成", "渡した"]:
        assert s in page, s
    # 測定履歴（前回比・初回比）
    for s in ["最新測定日", "履歴 2件", "初回 0.737", "太もも左右差", "初回比 -1.5cm", "2026/07/01", "68.5cm",
              "前回 <span class=\"bj-delta decrease\">-1.5cm</span>", "前回 <span class=\"bj-delta neutral\">-</span>", "順調"]:
        assert s in page, s
    # アンケート回答履歴（アンケートごと、最新が新しい順）
    assert page.find("<strong>施術後アンケート</strong>") < page.find("<strong>計測時アンケート</strong>")
    assert "回答数: 13件 / 最新: 2026/07/10 10:00" in page and "回数券 10回券 / 2枚目 / 3回目" in page and "ゴミ箱" not in page
    # 印が立っていないので保存ボタンは押せない
    assert 'disabled title="' + gate.断る文() in page


def test_回答しか無い方と何も無い方(as_owner):
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    page = as_owner.get(_url("山田太郎")).content.decode()
    assert "<h2>山田太郎</h2>" in page and "現在の回数券スタンプ情報はありません。" in page
    assert "まだ顧客メモはありません。" in page and "まだ測定履歴はありません。" in page and "フリガナ: -" in page
    assert ">未設定</span> 顧客側でまだ通知設定が同期されていません。" in page
    # 特典の設定が無い（既定）ときの文言
    assert "特典がまだ設定されていません。「🎁 特典」で節目と内容をご登録ください。" in page
    assert as_owner.get(_url("いない人")).status_code == 404
    # 表示をオフにするとマイルストーン特典の段は出ない
    _特典設定(milestoneRewardConfig={"enabled": False, "milestones": []})
    assert "マイルストーン特典" not in as_owner.get(_url("山田太郎")).content.decode()


# ═════════════════════════════════════════════════════════
# 書き込み
# ═════════════════════════════════════════════════════════

def test_印が立つまでは書けない(as_owner):
    _花子()
    r = as_owner.post(_url("佐藤花子", "memo"), {"at": "2026-09-10", "memo": "書けない"}, follow=True)
    assert gate.断る文() in r.content.decode()
    assert CustomerProfile.objects.get(name="佐藤花子").latest_memo == "新しい"
    for 経路 in ("save", "delete", "passcode", "stamp", "redemption", "link"):
        r = as_owner.post(_url("佐藤花子", 経路), {"name": "佐藤花子", "delta": "1", "threshold": "1"}, follow=True)
        assert gate.断る文() in r.content.decode(), 経路
    assert CustomerProfile.objects.filter(name="佐藤花子").exists() and Response.objects.count() == 15
    r = as_owner.post("/manage/bijiris/customers/add/", {"name": "新しい人"}, follow=True)
    assert gate.断る文() in r.content.decode() and not CustomerProfile.objects.filter(name="新しい人").exists()


def test_顧客情報の保存はお名前の変更が回答と測定に付いていき回数券は最新の回答も直す(as_owner):
    _花子()
    gate.切り替える("server")
    r = as_owner.post(_url("佐藤花子", "save"), {"name": "佐藤 花子", "member_number": "15", "ticket_plan": "6回券",
                                              "ticket_sheet": "3枚目", "ticket_round": "2回目"}, follow=True)
    html = r.content.decode()
    assert "顧客情報を保存しました。" in html and r.redirect_chain[-1][0] == _url("佐藤 花子")
    p = CustomerProfile.objects.get(name="佐藤 花子")
    assert p.member_number == "MYM-0015" and p.aliases == ["佐藤花子"] and p.admin_managed and p.member_id == ""
    assert p.active_ticket_card == {"plan": "6回券", "sheetNumber": 3, "round": 2} and p.active_ticket_card_source == "admin"
    assert Response.objects.filter(customer_name="佐藤花子").count() == 0 and Response.objects.filter(customer_name="佐藤 花子").count() == 15
    assert Measurement.objects.filter(customer_name="佐藤 花子", member_number="MYM-0015").count() == 2
    # 最新の回数券回答（ゴミ箱は除く）の答えが書き換わる。それより前は変わらない
    最新 = Response.objects.get(response_id="r2-3")
    assert 最新.answer_of("q_bijiris_session_ticket_plan") == "6回券" and 最新.answer_of("q_bijiris_session_ticket_sheet") == "3枚目"
    assert 最新.managed_at is not None and Response.objects.get(response_id="r2-2").answer_of("q_bijiris_session_ticket_plan") == "10回券"
    assert Response.objects.get(response_id="trash1").answer_of("q_bijiris_session_ticket_plan") == "10回券"
    # 言い分は旧アプリと同じ
    html = as_owner.post(_url("佐藤 花子", "save"), {"name": "", "ticket_plan": ""}, follow=True).content.decode()
    assert "お名前を入力してください。" in html
    html = as_owner.post(_url("佐藤 花子", "save"), {"name": "佐藤 花子", "ticket_plan": "6回券"}, follow=True).content.decode()
    assert "回数券情報は種類・何枚目・何回目をすべて選択してください。" in html
    # 回数券を空で保存しても、いまのカードは消えない（旧アプリの shouldUpdateTicket と同じ）
    as_owner.post(_url("佐藤 花子", "save"), {"name": "佐藤 花子"}, follow=True)
    assert CustomerProfile.objects.get(name="佐藤 花子").active_ticket_card["plan"] == "6回券"
    # 回答しか無い方を保存するとプロフィールができる
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    as_owner.post(_url("山田太郎", "save"), {"name": "山田太郎", "member_number": "MYM-0020"}, follow=True)
    assert CustomerProfile.objects.get(name="山田太郎").member_number == "MYM-0020"


def test_まゆみの会員と結ぶのは受付が選んだときだけ(as_owner):
    _花子()
    Member.objects.create(member_id="MYM-0012", name="佐藤花子")
    Member.objects.create(member_id="MYM-0099", name="別の人", deleted=True)
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "link"), {"member_id": "MYM-0012"}, follow=True).content.decode()
    assert "まゆみの会員 MYM-0012 と結びました。" in html and CustomerProfile.objects.get(name="佐藤花子").member_id == "MYM-0012"
    assert '<span class="badge badge-blue">まゆみの会員 MYM-0012</span>' in html
    html = as_owner.post(_url("佐藤花子", "link"), {"member_id": "MYM-0099"}, follow=True).content.decode()
    assert "その会員番号の会員が見つかりません。" in html and CustomerProfile.objects.get(name="佐藤花子").member_id == "MYM-0012"
    html = as_owner.post(_url("佐藤花子", "link"), {"member_id": ""}, follow=True).content.decode()
    assert "結びつきを外しました。" in html and CustomerProfile.objects.get(name="佐藤花子").member_id == ""
    # 一覧でも結んだ会員が添えられる
    as_owner.post(_url("佐藤花子", "link"), {"member_id": "MYM-0012"})
    assert "まゆみの会員 MYM-0012" in as_owner.get("/manage/bijiris/customers/").content.decode()


def test_顧客の削除は回答と写真と測定とプロフィールを全部消す(as_owner):
    _花子()
    ResponsePhoto.objects.create(response=Response.objects.get(response_id="r1-1"), question_id="q", drive_file_id="F")
    CustomerProfile.objects.create(name="別名を持つ人", aliases=["佐藤花子", "他"])
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "delete"), follow=True).content.decode()
    assert "顧客情報を削除しました。（回答 15件も削除）" in html
    assert not CustomerProfile.objects.filter(name="佐藤花子").exists() and Response.objects.count() == 1
    assert ResponsePhoto.objects.count() == 0 and Measurement.objects.count() == 0
    assert CustomerProfile.objects.get(name="別名を持つ人").aliases == ["他"]
    assert as_owner.post(_url("佐藤花子", "delete")).status_code == 404


def test_パスコードの再設定を許可すると今のパスコードが消えて30分の期限が付く(as_owner):
    p = _花子()
    p.passcode_hash, p.passcode_salt = "h", "s"
    p.save()
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "passcode"), follow=True).content.decode()
    assert "30分間、お客様がパスコードを設定できます。アプリの「パスコードを忘れた方」からお進みください。" in html
    p.refresh_from_db()
    assert not p.has_passcode and p.passcode_setup_until.endswith("Z")
    期限 = datetime.fromisoformat(p.passcode_setup_until.replace("Z", "+00:00"))
    assert 29 <= (期限 - timezone.now()).total_seconds() / 60 <= 30
    assert "パスコード未設定" in html and "再設定の許可期限" in html
    # プロフィールが無い方には許可できない
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    assert "お客様が見つかりませんでした。" in as_owner.post(_url("山田太郎", "passcode"), follow=True).content.decode()


def test_顧客メモは新しい順で同じ日の同じメモは重ねない(as_owner):
    _花子()
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "memo"), {"at": "2026-08-15", "memo": "間"}, follow=True).content.decode()
    assert "顧客メモを保存しました。" in html and "メモ履歴 3件" in html
    p = CustomerProfile.objects.get(name="佐藤花子")
    assert [e["memo"] for e in p.memo_entries] == ["新しい", "間", "古い"] and p.latest_memo == "新しい"
    as_owner.post(_url("佐藤花子", "memo"), {"at": "2026-08-15", "memo": "間"})
    assert len(CustomerProfile.objects.get(name="佐藤花子").memo_entries) == 3
    as_owner.post(_url("佐藤花子", "memo"), {"at": "2026-09-20", "memo": "いちばん新しい"})
    assert CustomerProfile.objects.get(name="佐藤花子").latest_memo == "いちばん新しい"
    assert "記録日を入力してください。" in as_owner.post(_url("佐藤花子", "memo"), {"at": "", "memo": "x"}, follow=True).content.decode()
    assert "メモを入力してください。" in as_owner.post(_url("佐藤花子", "memo"), {"at": "2026-09-01", "memo": " "}, follow=True).content.decode()
    # 回答しか無い方にもメモを持たせられる（プロフィールができる）
    _回答("山田太郎", 日本時間(2026, 8, 1, 9, 0), "y1")
    as_owner.post(_url("山田太郎", "memo"), {"at": "2026-09-01", "memo": "初めて"})
    assert CustomerProfile.objects.get(name="山田太郎").latest_memo == "初めて"


def test_スタンプの手当ては1個ずつで合計がマイナスにならない(as_owner):
    _花子()
    _特典設定()
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "stamp"), {"delta": "1"}, follow=True).content.decode()
    assert "スタンプを1個足しました。" in html and "スタンプ: <b>2個</b>（施術後アンケートから 0個 ＋ 手当て +2個）" in html
    as_owner.post(_url("佐藤花子", "stamp"), {"delta": "-1"})
    html = as_owner.post(_url("佐藤花子", "stamp"), {"delta": "-1"}, follow=True).content.decode()
    assert "スタンプを1個戻しました。" in html and CustomerProfile.objects.get(name="佐藤花子").ticket_stamp_adjustment == 0
    html = as_owner.post(_url("佐藤花子", "stamp"), {"delta": "-1"}, follow=True).content.decode()
    assert "これ以上は減らせません。" in html and CustomerProfile.objects.get(name="佐藤花子").ticket_stamp_adjustment == 0
    assert "1個ずつ" in as_owner.post(_url("佐藤花子", "stamp"), {"delta": "5"}, follow=True).content.decode()


def test_特典の受け取りは渡したの付け外し(as_owner):
    _花子()
    _特典設定()
    gate.切り替える("server")
    html = as_owner.post(_url("佐藤花子", "redemption"), {"threshold": "1", "handed": "1"}, follow=True).content.decode()
    assert "渡し済みにしました。" in html and "受取済み" in html and "（未受取はありません）" in html
    r = CustomerProfile.objects.get(name="佐藤花子").reward_redemptions
    assert r["1"]["handed"] is True and r["1"]["handedAt"].endswith("Z")
    html = as_owner.post(_url("佐藤花子", "redemption"), {"threshold": "1", "handed": "0"}, follow=True).content.decode()
    assert "未渡しに戻しました。" in html and "未受取 1件" in html
    assert CustomerProfile.objects.get(name="佐藤花子").reward_redemptions is None
    assert "しきい値が正しくありません。" in as_owner.post(_url("佐藤花子", "redemption"), {"threshold": "x"}, follow=True).content.decode()


def test_顧客の追加(as_owner):
    gate.切り替える("server")
    r = as_owner.post("/manage/bijiris/customers/add/", {"name": "新しい人", "name_kana": "アタラシイ ヒト", "member_number": "7"}, follow=True)
    assert "顧客を追加しました。" in r.content.decode() and r.redirect_chain[-1][0] == _url("新しい人")
    p = CustomerProfile.objects.get(name="新しい人")
    assert p.name_kana == "アタラシイヒト" and p.member_number == "MYM-0007" and p.admin_managed and p.member_id == ""
    assert "すでに登録されています。" in as_owner.post("/manage/bijiris/customers/add/", {"name": "新しい人"}, follow=True).content.decode()
    assert "お名前を入力してください。" in as_owner.post("/manage/bijiris/customers/add/", {"name": " "}, follow=True).content.decode()
