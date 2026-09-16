"""ビジリス管理「回答管理」（apps/bijiris/views_responses.py）と「集計」（views_dashboard.py）。
中身は旧ビジリス管理アプリ #page-responses / #page-dashboard と同じ。"""

import io
from datetime import datetime

import pytest
from django.utils import timezone

from apps.bijiris import answers, gate
from apps.bijiris.models import CustomerProfile, Response, ResponsePhoto, Survey
from apps.manage import images

pytestmark = pytest.mark.django_db

LIST = "/manage/bijiris/responses/"
DASH = "/manage/bijiris/"


def _q(qid, label, type_="text", **extra):
    q = {"id": qid, "label": label, "type": type_, "required": True, "options": [], "placeholder": "",
         "visibilityConditions": [], "visibleWhen": None}
    q.update(extra)
    return q


def _t(y, m, d, h=10):
    return timezone.make_aware(datetime(y, m, d, h, 5))


お悩み1 = answers.お悩みの分類[0]["options"][0]   # トイレ
お悩み2 = answers.お悩みの分類[1]["options"][0]   # お腹まわり


@pytest.fixture
def data():
    session = Survey.objects.create(survey_id="survey_session", title="施術後アンケート", description="施術後", sort_order=0, questions=[
        _q("q_bijiris_session_ticket_plan", "回数券", "choice", options=["6回券", "10回券"]),
        _q("q_bijiris_session_ticket_sheet", "何枚目", "choice", options=["1枚目", "2枚目"]),
        _q("q_bijiris_session_ticket_round", "何回目", "choice", options=["1回目", "2回目", "3回目"]),
        _q(answers.お悩みの設問ID, "お悩み", "checkbox", options=[お悩み1, お悩み2]),
        _q("q_rate", "満足度", "rating"),
        _q("q_free", "感想", "textarea", required=False),
    ])
    measure = Survey.objects.create(survey_id="survey_measure", title="計測時アンケート", description="計測", sort_order=1, questions=[
        _q("q_measure_timing", "計測のタイミング", "choice", options=["初回計測時", "回数券終了時"]),
        _q("q_measure_photos", "全身写真", "photo"),
        _q("q_measure_waist", "ウエスト"),
    ])
    CustomerProfile.objects.create(name="佐藤花子", member_number="MYM-0012", name_kana="サトウ ハナコ", client_ids=["client-1"])
    r1 = Response.objects.create(response_id="r1", survey=session, survey_key="survey_session", survey_title="施術後アンケート",
                                 submitted_at=_t(2026, 8, 1), client_id="client-1", customer_name="佐藤花子", status="new",
                                 answers=[
                                     {"questionId": "q_bijiris_session_ticket_plan", "label": "回数券", "type": "choice", "value": "10回券"},
                                     {"questionId": "q_bijiris_session_ticket_sheet", "label": "何枚目", "type": "choice", "value": "1枚目"},
                                     {"questionId": "q_bijiris_session_ticket_round", "label": "何回目", "type": "choice", "value": "2回目"},
                                     {"questionId": answers.お悩みの設問ID, "label": "お悩み", "type": "checkbox", "value": f"{お悩み1}, {お悩み2}"},
                                     {"questionId": "q_rate", "label": "満足度", "type": "rating", "value": "4"},
                                     {"questionId": "q_free", "label": "感想", "type": "textarea", "value": "よかった"},
                                 ])
    r2 = Response.objects.create(response_id="r2", survey=session, survey_key="survey_session", survey_title="施術後アンケート",
                                 submitted_at=_t(2026, 9, 1), client_id="client-1", customer_name="佐藤花子", status="done",
                                 admin_memo="電話済み", answers=[
                                     {"questionId": "q_bijiris_session_ticket_plan", "label": "回数券", "type": "choice", "value": "10回券"},
                                     {"questionId": "q_bijiris_session_ticket_round", "label": "何回目", "type": "choice", "value": "3回目"},
                                     {"questionId": answers.お悩みの設問ID, "label": "お悩み", "type": "checkbox", "value": お悩み1},
                                     {"questionId": "q_rate", "label": "満足度", "type": "rating", "value": "5"},
                                 ])
    r3 = Response.objects.create(response_id="r3", survey=measure, survey_key="survey_measure", survey_title="計測時アンケート",
                                 submitted_at=_t(2026, 8, 10), client_id="client-1", customer_name="佐藤花子", status="checked",
                                 answers=[{"questionId": "q_measure_timing", "label": "計測のタイミング", "type": "choice", "value": "初回計測時"},
                                          {"questionId": "q_measure_photos", "label": "全身写真", "type": "photo", "value": "", "files": [{"fileId": "A"}]},
                                          {"questionId": "q_measure_waist", "label": "ウエスト", "type": "text", "value": "70"}])
    ResponsePhoto.objects.create(response=r3, question_id="q_measure_photos", drive_file_id="A", name="写真01.jpg", kind="before",
                                 captured_at="2026-08-10 09:00")
    r4 = Response.objects.create(response_id="r4", survey=measure, survey_key="survey_measure", survey_title="計測時アンケート",
                                 submitted_at=_t(2026, 9, 5), client_id="client-1", customer_name="佐藤花子", status="new",
                                 answers=[{"questionId": "q_measure_timing", "label": "計測のタイミング", "type": "choice", "value": "回数券終了時"},
                                          {"questionId": "q_measure_photos", "label": "全身写真", "type": "photo", "value": "", "files": []},
                                          {"questionId": "q_measure_waist", "label": "ウエスト", "type": "text", "value": "66"}])
    url = images.保存する(io.BytesIO(_jpeg()), "bijiris")
    ResponsePhoto.objects.create(response=r4, question_id="q_measure_photos", drive_file_id="B", name="写真02.jpg", kind="after", url=url)
    ResponsePhoto.objects.create(response=r4, question_id="q_measure_photos", drive_file_id="C", name="写真03.jpg", kind="after", order=1)
    r5 = Response.objects.create(response_id="r5", survey=measure, survey_key="survey_measure", survey_title="計測時アンケート",
                                 submitted_at=_t(2026, 7, 1), customer_name="山田太郎", status="trash", admin_memo="取り消し",
                                 answers=[{"questionId": "q_measure_waist", "label": "ウエスト", "type": "text", "value": "80"}])
    return {"session": session, "measure": measure, "r1": r1, "r2": r2, "r3": r3, "r4": r4, "r5": r5}


def _jpeg():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (400, 300), (200, 100, 50)).save(buf, format="JPEG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 一覧（アンケートタイトル一覧 → 回答者の一覧）
# ---------------------------------------------------------------------------

def test_一覧の絞り込みの欄と文言(as_owner, data):
    page = as_owner.get(LIST).content.decode()
    assert "<h1>回答管理</h1>" in page
    for label in ("アンケート", "対応状況", "検索", "会員番号", "フリガナ", "管理メモ", "お悩みカテゴリ", "回数券", "何枚目", "何回目",
                  "写真", "写真枚数", "開始日", "終了日", "未対応のみ", "CSV出力", "PDF出力"):
        assert label in page, label
    for ph in ("会員番号・氏名・アンケート・回答内容", "会員番号で絞り込み", "ふりがなで絞り込み"):
        assert ph in page, ph
    for opt in ("すべてのカテゴリ", "【その他】", "すべての回数券", ">6回券<", ">10回券<", "すべての枚数", ">20枚目<", "すべての回数", ">10回目<",
                "メモあり", "メモなし", "写真あり", "写真なし", "1-2枚", "3-4枚", "5枚以上", ">ゴミ箱<"):
        assert opt in page, opt
    # アンケートタイトル一覧（ゴミ箱は数えない）。最新の回答が新しい順
    assert "アンケートタイトル一覧" in page and "アンケートを選ぶと、そのアンケートの過去回答を最新順で表示します。" in page
    assert "回答数: 2件 / 回答者: 1名" in page and "最新回答: 2026/09/05 10:05" in page and "最新回答: 2026/09/01 10:05" in page
    assert page.index("計測時アンケート</strong>") < page.index("施術後アンケート</strong>")
    assert "未対応 1件" in page and "未対応なし" not in page
    assert "読むだけ" in page and gate.断る文() in page


def test_絞り込みは旧アプリと同じ条件(as_owner, data):
    def 群(**params):
        page = as_owner.get(LIST, params).content.decode()
        return [t for t in ("施術後アンケート", "計測時アンケート") if f"{t}</strong>" in page], page

    assert 群(survey="survey_session")[0] == ["施術後アンケート"]
    assert 群(status="done")[0] == ["施術後アンケート"]
    # ゴミ箱は対応状況で選んだときだけ
    g, page = 群(status="trash")
    assert g == ["計測時アンケート"] and "回答数: 1件 / 回答者: 1名" in page
    assert 群(q="電話済み")[0] == ["施術後アンケート"]
    assert 群(q="mym-0012")[0] == ["施術後アンケート", "計測時アンケート"]
    assert 群(q="よかった")[0] == ["施術後アンケート"]
    assert 群(member="0012")[0] == ["施術後アンケート", "計測時アンケート"]
    assert 群(member="9999")[0] == []
    assert 群(kana="さとう")[0] == [] and 群(kana="サトウハナコ")[0] == ["施術後アンケート", "計測時アンケート"]
    assert 群(memo="with")[0] == ["施術後アンケート"]
    g, _ = 群(memo="without")
    assert g == ["施術後アンケート", "計測時アンケート"]
    assert 群(concern="belly")[0] == ["施術後アンケート"] and 群(concern="posture")[0] == []
    assert 群(plan="10回券")[0] == ["施術後アンケート"] and 群(plan="6回券")[0] == []
    assert 群(sheet="1枚目")[0] == ["施術後アンケート"] and 群(round="3回目")[0] == ["施術後アンケート"]
    assert 群(photo="with")[0] == ["計測時アンケート"] and 群(photo="without")[0] == ["施術後アンケート"]
    assert 群(photos="1-2")[0] == ["計測時アンケート"] and 群(photos="5plus")[0] == []
    assert 群(**{"from": "2026-09-01"})[0] == ["施術後アンケート", "計測時アンケート"]
    assert 群(**{"from": "2026-09-02"})[0] == ["計測時アンケート"]
    assert 群(to="2026-08-31")[0] == ["施術後アンケート", "計測時アンケート"] and 群(to="2026-07-31")[0] == []
    _, page = 群(q="どこにもない")
    assert "条件に一致するアンケートはありません。" in page
    # 「未対応のみ」は対応状況=未対応
    page = as_owner.get(LIST).content.decode()
    assert "status=new" in page


def test_アンケートを選ぶと回答者の一覧(as_owner, data):
    page = as_owner.get(LIST, {"group": "survey_measure"}).content.decode()
    assert "<h2>計測時アンケート</h2>" in page and "回答者名を押すと詳細を表示します。" in page and ">戻る<" in page
    # 最新順。ゴミ箱の r5 は出ない
    assert page.index("/responses/r4/") < page.index("/responses/r3/") and "/responses/r5/" not in page
    assert page.count("回答削除") == 2 and "完全削除" not in page
    page = as_owner.get(LIST, {"group": "survey_measure", "status": "trash"}).content.decode()
    assert "/responses/r5/" in page and "完全削除" in page and "この回答を完全削除しますか？" in page
    page = as_owner.get(LIST, {"group": "survey_measure", "status": "trash", "q": "無い"}).content.decode()
    # 条件に合う回答が無ければタイトル一覧に戻る
    assert "条件に一致するアンケートはありません。" in page


# ---------------------------------------------------------------------------
# 詳細
# ---------------------------------------------------------------------------

def test_詳細は旧アプリと同じ項目(as_owner, data):
    page = as_owner.get("/manage/bijiris/responses/r2/").content.decode()
    assert "<h1>施術後アンケート</h1>" in page and "MYM-0012 / 佐藤花子 / 2026/09/01 10:05" in page
    for b in ("印刷", "回答削除", ">戻る<", ">閉じる<", "ゴミ箱へ移動", "この回答をゴミ箱へ移動しますか？"):
        assert b in page, b
    assert "対応済み" in page
    # 回数券の印
    assert "回数券</span><span>10回券</span>" in page and "何回目</span><span>3回目</span>" in page
    # 前回比較: r1 との違いだけ
    assert "前回比較" in page and "前回回答: 2026/08/01 10:05" in page
    assert "前回: 1枚目" in page and "今回: 未回答" in page      # 何枚目が空になった
    assert "前回: 4" in page and "今回: 5" in page               # 満足度
    assert "前回: よかった" in page and "今回: 未回答" in page
    assert "前回: 10回券" not in page                            # 同じ値は出さない
    # 写真比較: このアンケートに写真は無い
    assert "写真比較" in page and "比較できる写真はありません。" in page
    # 回答の全項目（直せる）と管理の欄
    for label in ("回数券", "何枚目", "何回目", "お悩み", "満足度", "感想", "対応状況", "管理メモ", ">保存<"):
        assert label in page, label
    assert 'name="a-q_rate"' in page and '<option value="5" selected>5</option>' in page
    assert 'name="a-q_free"' in page and "選択してください" in page
    assert "【トイレ・デリケートなお悩み】" in page and "1件" in page and f'value="{お悩み1}"' in page
    assert ">電話済み</textarea>" in page and ">必須<" in page and ">任意<" in page
    assert as_owner.get("/manage/bijiris/responses/nothing/").status_code == 404


def test_詳細の写真と写真比較(as_owner, data):
    page = as_owner.get("/manage/bijiris/responses/r4/").content.decode()
    assert "アップロード写真" in page and "回答で添付された写真を表示しています。" in page
    assert "/media/bijiris/" in page and "写真02.jpg" in page
    assert "未取り込み" in page and "写真03.jpg" in page
    # 写真比較: 初回（r3）と今回（r4）
    assert "初回・前回・今回の写真を並べて確認できます。" in page
    assert "<strong>初回</strong>" in page and "<strong>今回</strong>" in page and "写真01.jpg" in page
    assert "撮影日: 2026-08-10 09:00" in page
    # 前回比較（r3）: ウエストと写真の枚数
    assert "前回: 70" in page and "今回: 66" in page and "前回: 1枚" in page and "今回: 2枚" in page
    # 写真の設問は入力欄にしない
    assert 'name="a-q_measure_photos"' not in page and 'name="a-q_measure_waist"' in page
    # r3 は写真があるので「初回 / 今回」
    page = as_owner.get("/manage/bijiris/responses/r3/").content.decode()
    assert "<strong>初回 / 今回</strong>" in page and "比較できる前回回答はありません。" in page
    # r1 は前回が無い
    page = as_owner.get("/manage/bijiris/responses/r1/").content.decode()
    assert "比較できる前回回答はありません。" in page


def test_ゴミ箱の回答の詳細(as_owner, data):
    page = as_owner.get("/manage/bijiris/responses/r5/").content.decode()
    assert page.count("完全削除") >= 2 and "この回答を完全削除しますか？" in page and "ゴミ箱へ移動" not in page
    assert "山田太郎 / 2026/07/01 10:05" in page and "MYM" not in page


# ---------------------------------------------------------------------------
# 書き込み
# ---------------------------------------------------------------------------

def test_印が立つまでは書けない(as_owner, data):
    r = as_owner.post("/manage/bijiris/responses/r1/save/", {"status": "done", "admin_memo": "x"}, follow=True)
    assert gate.断る文() in r.content.decode() and r.redirect_chain[-1][0] == "/manage/bijiris/responses/r1/"
    as_owner.post("/manage/bijiris/responses/r1/trash/", follow=True)
    as_owner.post("/manage/bijiris/responses/r5/purge/", follow=True)
    assert Response.objects.get(response_id="r1").status == "new" and Response.objects.filter(response_id="r5").exists()


def test_保存は対応状況とメモと回答を直して管理更新日時を進める(as_owner, data):
    gate.切り替える("server")
    r = as_owner.post("/manage/bijiris/responses/r1/save/", {
        "status": "checked", "admin_memo": " 確認した ", "back": "group=survey_session",
        "a-q_bijiris_session_ticket_plan": "6回券", "a-q_bijiris_session_ticket_sheet": "2枚目", "a-q_bijiris_session_ticket_round": "1回目",
        f"a-{answers.お悩みの設問ID}": ["", お悩み2], "a-q_rate": "3", "a-q_free": "",
    }, follow=True)
    page = r.content.decode()
    assert "回答管理を保存しました。" in page and r.redirect_chain[-1][0] == "/manage/bijiris/responses/r1/?back=group%3Dsurvey_session"
    res = Response.objects.get(response_id="r1")
    assert res.status == "checked" and res.admin_memo == "確認した" and res.managed_at is not None
    assert res.answer_of("q_bijiris_session_ticket_plan") == "6回券" and res.answer_of("q_rate") == "3" and res.answer_of("q_free") == ""
    assert res.answer_of(answers.お悩みの設問ID) == お悩み2
    assert [a["questionId"] for a in res.answers] == [q["id"] for q in data["session"].questions]
    assert res.answers[0]["label"] == "回数券" and res.answers[0]["type"] == "choice"


def test_保存の検査はCodegsと同じ(as_owner, data):
    gate.切り替える("server")
    r = as_owner.post("/manage/bijiris/responses/r1/save/", {"status": "new", "a-q_rate": "9"}, follow=True)
    assert "評価は1から5で回答してください。" in r.content.decode()
    r = as_owner.post("/manage/bijiris/responses/r1/save/", {"status": "new", "a-q_bijiris_session_ticket_plan": "100回券"}, follow=True)
    assert "選択肢から回答してください。" in r.content.decode()
    r = as_owner.post("/manage/bijiris/responses/r1/save/", {"status": "new", f"a-{answers.お悩みの設問ID}": ["知らない項目"]}, follow=True)
    assert "選択肢から回答してください。" in r.content.decode()
    assert Response.objects.get(response_id="r1").answer_of("q_rate") == "4"
    # 写真の設問はそのまま残る。知らない対応状況は「未対応」
    r = as_owner.post("/manage/bijiris/responses/r3/save/", {"status": "変", "a-q_measure_waist": "71"}, follow=True)
    res = Response.objects.get(response_id="r3")
    assert res.status == "new" and res.answer_of("q_measure_waist") == "71"
    assert res.answer_of("q_measure_photos") == "" and [a for a in res.answers if a["questionId"] == "q_measure_photos"][0]["files"] == [{"fileId": "A"}]


def test_表示条件で見えない設問は空になる(as_owner):
    gate.切り替える("server")
    s = Survey.objects.create(survey_id="s", title="条件付き", description="x", questions=[
        _q("k1", "回数券ですか", "choice", options=["はい", "いいえ"]),
        _q("k2", "何回目", "text", visibilityConditions=[{"questionId": "k1", "value": "はい"}], visibleWhen={"questionId": "k1", "value": "はい"}),
    ])
    r = Response.objects.create(response_id="c1", survey=s, survey_key="s", survey_title="条件付き", submitted_at=_t(2026, 9, 1),
                                customer_name="A", answers=[{"questionId": "k1", "value": "はい"}, {"questionId": "k2", "value": "3"}])
    as_owner.post("/manage/bijiris/responses/c1/save/", {"status": "new", "a-k1": "いいえ", "a-k2": "3"})
    r.refresh_from_db()
    assert r.answer_of("k1") == "いいえ" and r.answer_of("k2") == ""


def test_ゴミ箱へ移動と戻すと完全削除(as_owner, data, settings):
    gate.切り替える("server")
    r = as_owner.post("/manage/bijiris/responses/r4/trash/", {"back": "group=survey_measure"}, follow=True)
    assert "回答をゴミ箱へ移動しました。" in r.content.decode() and r.redirect_chain[-1][0] == "/manage/bijiris/responses/?group=survey_measure"
    res = Response.objects.get(response_id="r4")
    assert res.trashed and res.managed_at is not None
    # ゴミ箱の回答は一覧に出ない・集計にも数えない
    assert "計測時アンケート</strong>" in as_owner.get(LIST).content.decode()
    assert "回答数: 1件 / 回答者: 1名" in as_owner.get(LIST).content.decode()
    # 戻すのは対応状況を選んで保存（旧アプリと同じ）
    as_owner.post("/manage/bijiris/responses/r4/save/", {"status": "new", "a-q_measure_waist": "66"})
    assert not Response.objects.get(response_id="r4").trashed
    # ゴミ箱に入っていないものは完全削除できない
    r = as_owner.post("/manage/bijiris/responses/r4/purge/", follow=True)
    assert "先にゴミ箱へ移動してください。" in r.content.decode() and Response.objects.filter(response_id="r4").exists()
    # ゴミ箱に入れてから完全削除: 写真の記録と実体も消える
    as_owner.post("/manage/bijiris/responses/r4/trash/")
    名前 = ResponsePhoto.objects.get(drive_file_id="B").url.rsplit("/", 1)[1]
    assert (settings.MEDIA_ROOT / "bijiris" / 名前).is_file()
    r = as_owner.post("/manage/bijiris/responses/r4/purge/", follow=True)
    assert "回答を完全削除しました。" in r.content.decode()
    assert not Response.objects.filter(response_id="r4").exists() and not ResponsePhoto.objects.filter(drive_file_id="B").exists()
    assert not (settings.MEDIA_ROOT / "bijiris" / 名前).exists()
    assert as_owner.post("/manage/bijiris/responses/r4/purge/").status_code == 404


def test_スタッフは入れない(client, staff, data):
    client.force_login(staff)
    assert client.get(LIST).status_code == 403
    assert client.get("/manage/bijiris/responses/r1/").status_code == 403
    assert client.post("/manage/bijiris/responses/r1/trash/").status_code == 403


# ---------------------------------------------------------------------------
# CSV・印刷
# ---------------------------------------------------------------------------

def test_CSV出力は旧アプリと同じ列(as_owner, data):
    r = as_owner.get("/manage/bijiris/responses/export.csv", {"survey": "survey_session"})
    assert r.status_code == 200 and r["Content-Type"].startswith("text/csv")
    assert r["Content-Disposition"].startswith('attachment; filename="survey-responses-')
    本文 = r.content.decode("utf-8-sig")
    行 = 本文.splitlines()
    assert 行[0] == '"日時","会員番号","お客様","アンケート","回数券","何枚目","何回目","対応状況","管理メモ","写真枚数","回答"'
    assert 行[1].startswith('"2026/09/01 10:05","MYM-0012","佐藤花子","施術後アンケート","10回券","","3回目","対応済み","電話済み","0","')
    assert f"お悩み: 【トイレ・デリケートなお悩み】 {お悩み1}" in 行[1] and "満足度: 5" in 行[1]
    assert 行[2].startswith('"2026/08/01 10:05","MYM-0012","佐藤花子","施術後アンケート","10回券","1枚目","2回目","未対応","","0","')
    assert r.content.startswith("\ufeff".encode("utf-8"))
    # 写真は URL（無ければ名前）
    r = as_owner.get("/manage/bijiris/responses/export.csv", {"survey": "survey_measure"})
    本文 = r.content.decode("utf-8-sig")
    assert "全身写真: https://api.example.com/media/bijiris/" in 本文 and "写真03.jpg" in 本文 and '"2"' in 本文
    # 無ければ文言を出して一覧へ
    r = as_owner.get("/manage/bijiris/responses/export.csv", {"q": "無い"}, follow=True)
    assert "出力対象の回答がありません。" in r.content.decode()


def test_印刷用の画面(as_owner, data):
    page = as_owner.get("/manage/bijiris/responses/r1/print/").content.decode()
    assert "<h1>施術後アンケート</h1>" in page and "MYM-0012 / 佐藤花子 / 2026/08/01 10:05" in page
    assert "<strong>満足度</strong><div>4</div>" in page and "window.print()" in page
    page = as_owner.get("/manage/bijiris/responses/r3/print/").content.decode()
    assert "<strong>全身写真</strong><div>写真01.jpg</div>" in page
    page = as_owner.get("/manage/bijiris/responses/print/", {"survey": "survey_session"}).content.decode()
    assert "<h1>回答一覧</h1>" in page and "件数: 2件" in page and "対応状況: 対応済み" in page
    assert "回数券: 回数券 10回券 / 何回目 3回目" in page and "写真枚数: 0" in page and "<strong>管理メモ</strong><div>電話済み</div>" in page
    r = as_owner.get("/manage/bijiris/responses/print/", {"q": "無い"}, follow=True)
    assert "出力対象の回答がありません。" in r.content.decode()


# ---------------------------------------------------------------------------
# 集計
# ---------------------------------------------------------------------------

def test_集計は旧アプリと同じ数字(as_owner, data):
    page = as_owner.get(DASH).content.decode()
    assert "<h1>集計</h1>" in page
    for label, n in (("回答数", 4), ("未対応", 2), ("回答者数", 1), ("アンケート数", 2)):
        assert f'<div class="stat-label">{label}</div><div class="stat-value">{n}</div>' in page, label
    # 最新の回答（新しい順・ゴミ箱を除く）
    i = page.index("最新の回答")
    j = page.index("アンケート別回答数")
    最新 = page[i:j]
    assert 最新.count("MYM-0012 / 佐藤花子") == 4 and "山田太郎" not in 最新
    assert 最新.index("計測時アンケート / 2026/09/05 10:05") < 最新.index("施術後アンケート / 2026/09/01 10:05")
    # アンケート別回答数（アンケートの並び）
    k = page.index("質問別分析")
    別 = page[j:k]
    assert 別.index("施術後アンケート") < 別.index("計測時アンケート") and 別.count("2件") == 2
    # 質問別分析: 先頭のアンケート
    分析 = page[k:page.index("月別回答件数")]
    assert "アンケートごとに回答傾向を確認できます。" in 分析
    assert "<span>10回券</span><strong>2件</strong>" in 分析 and "<span>1枚目</span><strong>1件</strong>" in 分析
    assert "<span>5</span><strong>1件</strong>" in 分析 and "<span>4</span><strong>1件</strong>" in 分析
    assert "<span>入力あり</span><strong>1件</strong>" in 分析
    assert "【トイレ・デリケートなお悩み】</strong><span>2件</span>" in 分析 and "【お腹まわり・便通のお悩み】</strong><span>1件</span>" in 分析
    # 月別回答件数（日本時間の月）
    月別 = page[page.index("月別回答件数"):]
    assert "アンケートごとの月別件数です。" in 月別
    assert 月別.index("2026-08") < 月別.index("2026-09")
    assert "2026-07" not in 月別   # ゴミ箱だけの月は出ない


def test_集計でアンケートを選ぶ(as_owner, data):
    page = as_owner.get(DASH, {"survey": "survey_measure"}).content.decode()
    分析 = page[page.index("質問別分析"):page.index("月別回答件数")]
    assert "<span>写真添付あり</span><strong>2件</strong>" in 分析
    assert "<span>初回計測時</span><strong>1件</strong>" in 分析 and "<span>回数券終了時</span><strong>1件</strong>" in 分析
    assert "<span>入力あり</span><strong>2件</strong>" in 分析
    assert '<option value="survey_measure" selected>' in page
    # 誰も答えていない選択式は「未回答 0件」
    Survey.objects.create(survey_id="s0", title="空", description="x", sort_order=9, questions=[_q("z", "選ぶ", "choice", options=["a", "b"])])
    page = as_owner.get(DASH, {"survey": "s0"}).content.decode()
    assert "<span>未回答</span><strong>0件</strong>" in page


def test_集計が空のときの文言(as_owner):
    page = as_owner.get(DASH).content.decode()
    for s in ("まだ回答はありません。", "アンケートはありません。", "アンケートを選択してください。", "まだ集計できる回答がありません。"):
        assert s in page, s
    for label in ("回答数", "未対応", "回答者数", "アンケート数"):
        assert f'<div class="stat-label">{label}</div><div class="stat-value">0</div>' in page
