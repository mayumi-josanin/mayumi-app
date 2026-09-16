"""ビジリス管理「アンケート管理」（apps/bijiris/views_surveys.py）。中身は旧ビジリス管理アプリ #page-surveys と同じ。"""

import json

import pytest
from django.utils import timezone

from apps.bijiris import gate
from apps.bijiris.models import Survey
from apps.bijiris.views_surveys import アンケートを整える

pytestmark = pytest.mark.django_db

LIST = "/manage/bijiris/surveys/"
NEW = "/manage/bijiris/surveys/new/"


def _q(qid, label, type_="text", **extra):
    q = {"id": qid, "label": label, "type": type_, "required": True, "options": [], "placeholder": "",
         "visibilityConditions": [], "visibleWhen": None}
    q.update(extra)
    return q


@pytest.fixture
def surveys():
    a = Survey.objects.create(survey_id="survey_a", title="施術後アンケート", description="施術後", status="published", sort_order=0,
                              accepting_responses=True, updated_at=timezone.now(),
                              questions=[_q("q1", "施術内容", "choice", options=["回数券", "都度"]), _q("q2", "感想", "textarea")])
    b = Survey.objects.create(survey_id="survey_b", title="計測時アンケート", description="計測", status="draft", sort_order=1,
                              accepting_responses=False, questions=[_q("p1", "全身写真", "photo")])
    return a, b


def _投稿(title="新しいアンケート", description="説明", questions=None, **extra):
    d = {"title": title, "description": description, "introMessage": "", "completionMessage": "", "status": "published",
         "acceptingResponses": "1", "startAt": "", "endAt": "",
         "questions": json.dumps(questions if questions is not None else [
             {"id": "", "label": "お名前", "type": "text", "required": True, "options": [], "visibilityConditions": []}])}
    d.update(extra)
    return d


# ---------------------------------------------------------------------------
# 一覧
# ---------------------------------------------------------------------------

def test_一覧は旧アプリと同じ項目(as_owner, surveys):
    page = as_owner.get(LIST).content.decode()
    assert "<h1>アンケート管理</h1>" in page and "新規作成" in page
    assert "アンケート一覧" in page and "新規作成、編集、削除ができます。" in page
    # 「N. タイトル」・公開状態・設問数・更新・受付
    assert "1. 施術後アンケート" in page and "2. 計測時アンケート" in page
    assert page.index("1. 施術後アンケート") < page.index("2. 計測時アンケート")
    assert "公開中" in page and "下書き" in page and "2問" in page and "1問" in page
    assert "受付中" in page and "受付停止中" in page and "更新: -" in page
    for b in ("↑", "↓", "複製", "アーカイブ"):
        assert b in page, b
    # 印が立っていないので「読むだけ」
    assert "読むだけ" in page and gate.断る文() in page
    assert 'class="active">アンケート管理</a>' in page


def test_一覧が空のときの文言(as_owner):
    assert "アンケートはまだありません。" in as_owner.get(LIST).content.decode()


def test_スタッフは入れない(client, staff):
    client.force_login(staff)
    assert client.get(LIST).status_code == 403
    assert client.post(NEW, _投稿()).status_code == 403


# ---------------------------------------------------------------------------
# 編集画面
# ---------------------------------------------------------------------------

def test_新規作成の画面に旧アプリと同じ入力欄(as_owner):
    page = as_owner.get(NEW).content.decode()
    assert "<h1>アンケート新規作成</h1>" in page and "タイトル、説明、質問項目を入力してください。" in page
    for label in ("タイトル", "説明", "回答前メッセージ", "回答完了メッセージ", "公開状態", "回答受付中", "受付開始日時", "受付終了日時", "質問項目", "質問を追加"):
        assert label in page, label
    for opt in (">公開<", ">下書き<", ">アーカイブ<"):
        assert opt in page, opt
    # 設問の入力欄（JS が組み立てる文言）と作成ボタン
    for s in ("質問文", "質問形式", "表示されたときに必須回答", "表示条件を追加", "選択肢を追加", "カテゴリは固定表示です",
              "テキスト", "長文", "単一選択", "複数選択", "5段階評価", "写真", ">作成<"):
        assert s in page, s
    # 新規は空の設問が1つ
    初期 = json.loads(page[page.index('id="initialQuestions"'):].split(">", 1)[1].split("</script>")[0])
    assert len(初期) == 1 and 初期[0]["label"] == "" and 初期[0]["type"] == "text"


def test_編集の画面はいまの中身が入る(as_owner, surveys):
    a, _ = surveys
    page = as_owner.get(f"/manage/bijiris/surveys/{a.survey_id}/edit/").content.decode()
    assert "<h1>アンケート編集</h1>" in page and "作成: -" in page and "更新: " in page
    assert 'value="施術後アンケート"' in page and ">施術後</textarea>" in page and ">保存<" in page
    for b in ("複製", "アーカイブ", "削除"):
        assert b in page, b
    初期 = json.loads(page[page.index('id="initialQuestions"'):].split(">", 1)[1].split("</script>")[0])
    assert [q["label"] for q in 初期] == ["施術内容", "感想"] and 初期[0]["options"] == ["回数券", "都度"]
    assert as_owner.get("/manage/bijiris/surveys/nothing/edit/").status_code == 404


# ---------------------------------------------------------------------------
# 保存の規則（validateSurveyPayload_）
# ---------------------------------------------------------------------------

def test_保存の規則はCodegsと同じ():
    値 = アンケートを整える({"title": " 題 ", "description": "説明", "questions": [
        {"label": "選ぶ", "type": "choice", "options": ["a", "", "b"], "visibleWhen": {"questionId": "x", "value": "y"}},
        {"label": "写真", "type": "photo", "required": False, "options": ["捨てる"]},
        {"label": "変な種類", "type": "unknown"},
    ]}, None)
    assert 値["title"] == "題" and 値["intro_message"] == "説明" and 値["completion_message"] == "ご回答ありがとうございました。"
    assert 値["status"] == "published" and 値["accepting_responses"] and 値["start_at"] is None
    q = 値["questions"]
    assert q[0]["id"].startswith("question_") and q[0]["options"] == ["a", "b"]
    assert q[0]["visibilityConditions"] == [{"questionId": "x", "value": "y"}] and q[0]["visibleWhen"] == {"questionId": "x", "value": "y"}
    assert q[1]["required"] is False and q[1]["options"] == [] and q[1]["visibleWhen"] is None
    assert q[2]["type"] == "text"
    for payload, 文 in [
        ({"title": "", "description": "a", "questions": [{"label": "x"}]}, "タイトルを入力してください。"),
        ({"title": "a", "description": "", "questions": [{"label": "x"}]}, "説明文を入力してください。"),
        ({"title": "a", "description": "a", "questions": []}, "質問は1つ以上必要です。"),
        ({"title": "a", "description": "a", "questions": [{"label": ""}]}, "質問文を入力してください。"),
        ({"title": "a", "description": "a", "questions": [{"label": "x", "type": "checkbox", "options": ["1"]}]}, "選択式の質問は選択肢を2つ以上入力してください。"),
        ({"title": "a", "description": "a", "startAt": "2026-09-20T10:00", "endAt": "2026-09-19T10:00", "questions": [{"label": "x"}]}, "受付終了日時は開始日時以降にしてください。"),
        ({"title": "a", "description": "a", "startAt": "いつか", "questions": [{"label": "x"}]}, "日時の形式が正しくありません。"),
    ]:
        with pytest.raises(ValueError, match=文):
            アンケートを整える(payload, None)


def test_印が立つまでは書けない(as_owner, surveys):
    a, _ = surveys
    r = as_owner.post(NEW, _投稿(), follow=True)
    assert r.redirect_chain[-1][0] == LIST and gate.断る文() in r.content.decode()
    assert Survey.objects.count() == 2
    for url in (f"/manage/bijiris/surveys/{a.survey_id}/edit/", f"/manage/bijiris/surveys/{a.survey_id}/move/",
                f"/manage/bijiris/surveys/{a.survey_id}/duplicate/", f"/manage/bijiris/surveys/{a.survey_id}/archive/",
                f"/manage/bijiris/surveys/{a.survey_id}/delete/"):
        r = as_owner.post(url, _投稿(), follow=True)
        assert gate.断る文() in r.content.decode(), url
    a.refresh_from_db()
    assert a.status == "published" and Survey.objects.count() == 2


def test_新規作成は最後に付いて公開の文言(as_owner, surveys):
    gate.切り替える("server")
    r = as_owner.post(NEW, _投稿(startAt="2026-09-20T10:00", endAt="2026-09-30T18:00"), follow=True)
    page = r.content.decode()
    assert "アンケートを公開しました。" in page and "<h1>アンケート編集</h1>" in page
    s = Survey.objects.get(title="新しいアンケート")
    assert s.survey_id.startswith("survey_") and s.sort_order == 2 and s.intro_message == "説明"
    assert s.completion_message == "ご回答ありがとうございました。" and s.created_at and s.updated_at
    assert timezone.localtime(s.start_at).strftime("%Y-%m-%d %H:%M") == "2026-09-20 10:00"
    assert s.questions[0]["label"] == "お名前" and s.questions[0]["id"].startswith("question_")
    # 下書きで作ると文言が変わる
    r = as_owner.post(NEW, _投稿(title="下書きの分", status="draft"), follow=True)
    assert "アンケートを作成しました。" in r.content.decode()
    assert Survey.objects.get(title="下書きの分").sort_order == 3


def test_入力の不備は同じ画面に戻って文言が出る(as_owner):
    gate.切り替える("server")
    r = as_owner.post(NEW, _投稿(title=""))
    page = r.content.decode()
    assert r.status_code == 200 and "タイトルを入力してください。" in page and "<h1>アンケート新規作成</h1>" in page
    assert Survey.objects.count() == 0
    r = as_owner.post(NEW, _投稿(questions=[{"label": "x", "type": "choice", "options": ["1"]}]))
    assert "選択式の質問は選択肢を2つ以上入力してください。" in r.content.decode()
    # 入力した設問は画面に残る（JS に渡す初期値）
    assert '"label": "x"' in r.content.decode()


def test_編集の保存は並びと作成日時を保つ(as_owner, surveys):
    gate.切り替える("server")
    a, _ = surveys
    Survey.objects.filter(pk=a.pk).update(created_at=timezone.now() - timezone.timedelta(days=10))
    a.refresh_from_db()
    r = as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/edit/", _投稿(
        title="施術後アンケート（改）", description="新しい説明", status="draft", acceptingResponses="",
        questions=[{"id": "q1", "label": "施術内容", "type": "choice", "options": ["回数券", "都度", "体験"], "required": False,
                    "visibilityConditions": []},
                   {"id": "", "label": "満足度", "type": "rating", "visibilityConditions": [{"questionId": "q1", "value": "回数券"}]}]),
        follow=True)
    assert "アンケートを保存しました。" in r.content.decode()
    a.refresh_from_db()
    assert a.title == "施術後アンケート（改）" and a.status == "draft" and not a.accepting_responses and a.sort_order == 0
    assert a.created_at < a.updated_at and (timezone.now() - a.created_at).days >= 9
    assert a.questions[0]["options"] == ["回数券", "都度", "体験"] and a.questions[0]["required"] is False
    assert a.questions[1]["visibleWhen"] == {"questionId": "q1", "value": "回数券"} and a.questions[1]["type"] == "rating"
    r = as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/edit/", _投稿(title="公開に戻す"), follow=True)
    assert "アンケートの公開内容を保存しました。" in r.content.decode()


def test_並べ替え(as_owner, surveys):
    gate.切り替える("server")
    a, b = surveys

    def 並び():
        return list(Survey.objects.order_by("sort_order").values_list("survey_id", flat=True))

    as_owner.post(f"/manage/bijiris/surveys/{b.survey_id}/move/", {"direction": "-1"})
    assert 並び() == ["survey_b", "survey_a"]
    # 端ではそのまま
    as_owner.post(f"/manage/bijiris/surveys/{b.survey_id}/move/", {"direction": "-1"})
    assert 並び() == ["survey_b", "survey_a"]
    as_owner.post(f"/manage/bijiris/surveys/{b.survey_id}/move/", {"direction": "1"})
    assert 並び() == ["survey_a", "survey_b"]
    assert [s.sort_order for s in Survey.objects.order_by("sort_order")] == [0, 1]


def test_複製とアーカイブと削除(as_owner, surveys):
    gate.切り替える("server")
    a, b = surveys
    r = as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/duplicate/", follow=True)
    assert "アンケートを複製しました。" in r.content.decode()
    c = Survey.objects.get(title="施術後アンケート（複製）")
    assert c.status == "draft" and c.sort_order == 2 and c.questions[0]["label"] == "施術内容" and c.survey_id != a.survey_id
    assert "<h1>アンケート編集</h1>" in r.content.decode()

    r = as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/archive/", follow=True)
    assert "アンケートをアーカイブしました。" in r.content.decode()
    a.refresh_from_db()
    assert a.status == "archived" and a.title == "施術後アンケート" and a.updated_at

    r = as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/delete/", follow=True)
    assert "アンケートを削除しました。" in r.content.decode() and r.redirect_chain[-1][0] == LIST
    assert not Survey.objects.filter(pk=a.pk).exists()
    assert [s.sort_order for s in Survey.objects.order_by("sort_order")] == [0, 1]
    assert as_owner.post(f"/manage/bijiris/surveys/{a.survey_id}/delete/").status_code == 404
