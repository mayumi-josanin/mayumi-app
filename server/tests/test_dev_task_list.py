"""タスク一覧（/manage/dev/tasks/）。院長の依頼 2026-09-18。

今までタスクはプロジェクト詳細の中にしか無く、「いま何が残っているか」を見るのに
プロジェクトを1つずつ開くことになっていた。**プロジェクトをまたいだ一覧**がこの画面。

ここで確かめるのは、院長が見たときに困らないこと。
既定で終わったタスクが隠れること、絞り込み・並べ替えが住所に残ること、
期限切れが目に入ること、一覧から直にステータスを進められること。
"""

import datetime

import pytest
from django.utils import timezone

from apps.dev.models import DevProject, DevTask

pytestmark = pytest.mark.django_db

URL = "/manage/dev/tasks/"


@pytest.fixture
def 今日():
    return timezone.localdate()


@pytest.fixture
def データ(owner, staff, 今日):
    """2つのプロジェクトに、状態も優先度もばらばらのタスクを置く。"""
    予約 = DevProject.objects.create(name="予約システム", status="in_progress")
    アプリ = DevProject.objects.create(name="お客様アプリ", status="in_progress")
    # 作る順にも意味がある（「新しい順」で t2 より t1 が先に来るように t2 を先に作る）
    t2 = DevTask.objects.create(project=予約, title="画面の文言を直す", status="in_progress", priority="low",
                                category="docs", due_date=今日)
    t1 = DevTask.objects.create(project=予約, title="二重予約を止める", status="open", priority="critical",
                                category="bug", assignee=owner, due_date=今日 + datetime.timedelta(days=5),
                                estimate_hours=3, actual_hours=1)
    t3 = DevTask.objects.create(project=予約, title="終わった仕事", status="done", priority="high")
    t4 = DevTask.objects.create(project=アプリ, title="会員の引っ越し", status="review", priority="medium",
                                category="infra", assignee=staff, due_date=今日 - datetime.timedelta(days=3))
    return {"予約": 予約, "アプリ": アプリ, "t1": t1, "t2": t2, "t3": t3, "t4": t4}


def _page(client, query=""):
    res = client.get(URL + query)
    assert res.status_code == 200
    return res.content.decode()


def _並び(page, データ):
    """画面に出ているタスクを、出てきた順に返す。"""
    出た = [(page.find(f'#{t.pk} {t.title}'), t.title)
            for t in [データ["t1"], データ["t2"], データ["t3"], データ["t4"]] if f'#{t.pk} {t.title}' in page]
    return [title for _, title in sorted(出た)]


# ========== 入れる人 ==========


def test_スタッフは入れない(client, staff):
    client.force_login(staff)
    assert client.get(URL).status_code == 403


def _tabs(page):
    i = page.find('<nav class="page-tabs"')
    return page[i:page.find("</nav>", i)] if i != -1 else ""


def test_タブは3つになる(as_owner, データ):
    """開発管理のタブは プロジェクト／タスク一覧／目安箱 の3つ。開いているタブが光る。"""
    tabs = _tabs(_page(as_owner))
    for ラベル in ["プロジェクト", "タスク一覧", "目安箱"]:
        assert ラベル in tabs
    assert 'class="active">タスク一覧</a>' in tabs
    # タスクの詳細を開いても「タスク一覧」のタブが光る
    詳細 = as_owner.get(f'/manage/dev/tasks/{データ["t1"].pk}/').content.decode()
    assert 'class="active">タスク一覧</a>' in _tabs(詳細)


# ========== 出るもの ==========


def test_プロジェクトごとにまとまって出る(as_owner, データ):
    page = _page(as_owner)
    # 見出しはプロジェクト名（絞り込みの選択肢と混ざらないよう、見出しの形で探す）
    assert "<a href=\"/manage/dev/{}/\">予約システム</a>".format(データ["予約"].pk) in page
    assert "お客様アプリ</a>" in page
    # 既定（優先度の高い順）は 緊急 のある予約システムが先
    assert page.find(">予約システム</a>") < page.find(">お客様アプリ</a>")
    # 見出しの残り件数（予約システムは完了を除いて2件）
    assert "残り 2件" in page and "残り 1件" in page
    # 期限切れはプロジェクトの見出しにも件数を出す
    assert "期限切れ 1件" in page


def test_既定では完了を隠し完了も見るで出す(as_owner, データ):
    page = _page(as_owner)
    assert "終わった仕事" not in page
    assert "残っているタスクはありません" not in page
    ぜんぶ = _page(as_owner, "?done=1")
    assert "終わった仕事" in ぜんぶ


def test_集計の札(as_owner, データ):
    page = _page(as_owner)
    札 = page.split('class="stats"')[1].split("<!-- 絞り込み -->")[0]
    for ラベル in ["残り", "作業中", "期限切れ", "今週が期限"]:
        assert ラベル in 札
    # 残り3・作業中1・期限切れ1・今週が期限1（t2 は今日が期限）
    数 = [s for s in 札.split("stat-value")]
    assert ">3<" in 数[1] and ">1<" in 数[2] and ">1<" in 数[3] and ">1<" in 数[4]


def test_期限切れの印(as_owner, データ):
    page = _page(as_owner)
    assert "（期限切れ）" in page
    assert "task-due-over" in page
    # 今日・明日は別の色で目立たせる
    assert "task-due-soon" in page


def test_空のときの文言(as_owner, owner):
    assert "残っているタスクはありません" in _page(as_owner)


# ========== 絞り込み ==========


def test_絞り込み_プロジェクト(as_owner, データ):
    page = _page(as_owner, f'?project={データ["予約"].pk}')
    assert "二重予約を止める" in page and "会員の引っ越し" not in page


def test_絞り込み_ステータス(as_owner, データ):
    page = _page(as_owner, "?status=in_progress")
    assert "画面の文言を直す" in page and "二重予約を止める" not in page


def test_絞り込み_優先度(as_owner, データ):
    page = _page(as_owner, "?priority=critical")
    assert "二重予約を止める" in page and "会員の引っ越し" not in page


def test_絞り込み_カテゴリ(as_owner, データ):
    page = _page(as_owner, "?category=infra")
    assert "会員の引っ越し" in page and "二重予約を止める" not in page


def test_絞り込み_担当(as_owner, データ, staff):
    page = _page(as_owner, f"?assignee={staff.pk}")
    assert "会員の引っ越し" in page and "二重予約を止める" not in page


def test_絞り込み_期限切れだけ(as_owner, データ):
    page = _page(as_owner, "?overdue=1")
    assert "会員の引っ越し" in page
    assert "二重予約を止める" not in page and "画面の文言を直す" not in page


def test_絞り込みは押した状態が分かる(as_owner, データ):
    page = _page(as_owner, "?status=review&overdue=1")
    assert '<option value="review" selected>' in page
    assert 'name="overdue" value="1" checked' in page


# ========== 並べ替え ==========


def test_並べ替え(as_owner, データ):
    # 既定は優先度の高い順（緊急→低／中）。プロジェクトごとにまとまるので、
    # 一番強い札を持つ予約システムのまとまりが先に出る
    assert _並び(_page(as_owner), データ) == ["二重予約を止める", "画面の文言を直す", "会員の引っ越し"]
    # 期限の近い順（期限切れの会員の引っ越しが先。まとまりの順も入れ替わる）
    page = _page(as_owner, "?sort=due")
    assert _並び(page, データ) == ["会員の引っ越し", "画面の文言を直す", "二重予約を止める"]
    assert page.find(">お客様アプリ</a>") < page.find(">予約システム</a>")
    # 新しい順（あとから作ったものが先）
    assert _並び(_page(as_owner, "?sort=new"), データ) == ["会員の引っ越し", "二重予約を止める", "画面の文言を直す"]
    # 知らない並べ替えを渡されても既定に戻すだけ（落とさない）
    assert _並び(_page(as_owner, "?sort=でたらめ"), データ)[0] == "二重予約を止める"


# ========== 一覧から進める ==========


def test_行からステータスを進められる(as_owner, データ):
    戻り先 = URL + "?status=open&sort=due"
    res = as_owner.post(f'/manage/dev/tasks/{データ["t1"].pk}/move/', {"status": "in_progress", "next": 戻り先})
    # 押したら同じ絞り込みのまま戻る
    assert res.status_code == 302 and res["Location"] == 戻り先
    データ["t1"].refresh_from_db()
    assert データ["t1"].status == "in_progress"


def test_よその住所へは戻さない(as_owner, データ):
    """next は同じ管理画面の中だけ。外の住所を入れられても、今までどおり詳細へ戻す。"""
    res = as_owner.post(f'/manage/dev/tasks/{データ["t1"].pk}/move/',
                        {"status": "review", "next": "https://example.com/"})
    assert res["Location"] == f'/manage/dev/{データ["予約"].pk}/'
