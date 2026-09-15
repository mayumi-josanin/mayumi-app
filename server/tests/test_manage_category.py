"""カテゴリ管理。**中身は旧管理アプリ（カテゴリ管理）と同じ**
（分類3つの追加・種別ごとの3つの表・使用数・行内の編集・使用中でも消せる）。
"""

import re
from datetime import date, datetime, timezone as tz

import pytest

from apps.content.models import CalendarEvent, Category, Menu, News, Product
from apps.manage.views_category import _通知カテゴリの割り当て

pytestmark = pytest.mark.django_db


@pytest.fixture
def cats(db):
    Category.objects.create(sheet_row=2, name="イベント情報", kind="お知らせ")
    Category.objects.create(sheet_row=3, name="まゆみのつぶやき", kind="ブログ")
    Category.objects.create(sheet_row=4, name="教室", kind="メニュー")
    Category.objects.create(sheet_row=5, name="NEWS", kind="通知")


def _表の中身(page, heading):
    """見出しから次の見出し（か終わり）までを切り出す。"""
    m = re.search(re.escape(heading) + r"(.*?)(<h3|{% endblock %}|$)", page, re.S)
    return m.group(1)


# ── 一覧の形 ─────────────────────────────────────────────


def test_追加の分類は旧管理アプリと同じ3つを同じ順で(as_owner, cats):
    page = as_owner.get("/manage/categories/").content.decode()
    options = re.findall(r'<option value="([^"]+)">([^<]+)</option>', page.split("登録済みカテゴリ一覧")[0])
    assert options == [("ブログ", "NEWS"), ("メニュー", "メニュー一覧"), ("通知", "お知らせ一覧（通知）")]
    assert 'placeholder="例：健康情報"' in page
    assert "➕ 新規カテゴリ追加" in page and "📋 登録済みカテゴリ一覧" in page
    assert "※各カテゴリはそれぞれのセクションで分類・フィルタとして使用されます。" in page


def test_種別ごとに3つの表に分かれる_お知らせ種別はNEWSの表(as_owner, cats):
    page = as_owner.get("/manage/categories/").content.decode()
    news = _表の中身(page, "📰 NEWS用カテゴリ")
    menu = _表の中身(page, "🛒 メニュー一覧用カテゴリ")
    notice = _表の中身(page, "📢 お知らせ一覧（通知）用カテゴリ")
    # 行内の入力（class="form-control" value="…"）で、どの表に載っているかを見る
    def 載る(表, 名):
        return f'class="form-control" value="{名}"' in 表

    assert 載る(news, "イベント情報") and 載る(news, "まゆみのつぶやき") and not 載る(news, "教室")
    assert 載る(menu, "教室") and not 載る(menu, "NEWS")
    assert 載る(notice, "NEWS") and not 載る(notice, "教室")
    for 表 in (news, menu, notice):
        assert "<th>カテゴリ名</th>" in 表 and "使用数" in 表 and "操作" in 表


def test_空の表には旧管理アプリと同じ文言(as_owner):
    page = as_owner.get("/manage/categories/").content.decode()
    assert page.count("カテゴリがありません") == 3


def test_同じ名前が2行あっても一覧には1つだけ(as_owner, cats):
    Category.objects.create(sheet_row=6, name="教室", kind="メニュー")
    page = as_owner.get("/manage/categories/").content.decode()
    assert page.count('class="form-control" value="教室"') == 1


# ── 使用数 ─────────────────────────────────────────────


def _使用数(page, name):
    m = re.search(r'value="' + re.escape(name) + r'"[^>]*>.*?<td[^>]*>(\d+)</td>', page, re.S)
    return int(m.group(1))


def test_使用数はNEWSとメニューをカテゴリ名で数える_消した分は数えない(as_owner, cats):
    News.objects.create(sheet_row=5, title="a", category="イベント情報", published=True)
    News.objects.create(sheet_row=6, title="b", category="イベント情報", published=False)
    News.objects.create(sheet_row=7, title="c", category="イベント情報", published=True, deleted=True)
    News.objects.create(sheet_row=8, title="", category="イベント情報", published=True)
    Menu.objects.create(sheet_row=2, name="m1", category="教室", published=True)
    Menu.objects.create(sheet_row=3, name="m2", category="教室", published=False)
    Menu.objects.create(sheet_row=4, name="m3", category="教室", published=True, deleted=True)
    page = as_owner.get("/manage/categories/").content.decode()
    assert _使用数(page, "イベント情報") == 2
    assert _使用数(page, "教室") == 2
    assert _使用数(page, "まゆみのつぶやき") == 0


def test_カテゴリ空のNEWSはお知らせとして数える(as_owner, cats):
    Category.objects.create(sheet_row=6, name="お知らせ", kind="お知らせ")
    News.objects.create(sheet_row=5, title="a", category="", published=True)
    page = as_owner.get("/manage/categories/").content.decode()
    assert _使用数(page, "お知らせ") == 1


def test_通知用カテゴリはお知らせ一覧に並ぶ件数を出どころごとに数える(as_owner, cats):
    Category.objects.create(sheet_row=6, name="ショップ", kind="通知")
    Category.objects.create(sheet_row=7, name="カレンダー", kind="通知")
    Category.objects.create(sheet_row=8, name="ホーム", kind="通知")
    新しい = datetime(2026, 5, 1, 9, 0, tzinfo=tz.utc)
    古い = datetime(2026, 3, 1, 9, 0, tzinfo=tz.utc)
    # NEWS: 公開・一覧から外していない・2026-04-01 以降 → 載る
    News.objects.create(sheet_row=5, title="a", category="x", published=True, updated_at=新しい)
    News.objects.create(sheet_row=6, title="b", category="x", published=True, updated_at=古い)  # 古い
    News.objects.create(sheet_row=7, title="c", category="x", published=False, updated_at=新しい)  # 非公開
    News.objects.create(sheet_row=8, title="d", category="x", published=True, updated_at=新しい,
                        notice_delisted_at=新しい)  # 一覧から外した
    News.objects.create(sheet_row=9, title="e", category="x", published=True)  # 日付なし → 載る
    # カレンダー: 休診と往診は並べない
    CalendarEvent.objects.create(sheet_row=2, title="教室", published=True, event_on=date(2026, 6, 1))
    CalendarEvent.objects.create(sheet_row=3, title="休診日", published=True, event_on=date(2026, 6, 2))
    CalendarEvent.objects.create(sheet_row=4, title="往診", published=True, event_on=date(2026, 6, 3))
    CalendarEvent.objects.create(sheet_row=5, title="お休みの話", category="イベント", published=True,
                                 event_on=date(2026, 6, 4))  # 区分があれば文面から当てない
    # 商品: 登録日の予備が無い。updated_at が無ければ載る
    Product.objects.create(sheet_row=2, name="p1", published=True)
    Product.objects.create(sheet_row=3, name="p2", published=True, updated_at=古い)
    # メニュー
    Menu.objects.create(sheet_row=2, name="m1", published=True, registered_on=date(2026, 4, 1))
    Menu.objects.create(sheet_row=3, name="m2", published=True, registered_on=date(2026, 3, 31))
    page = as_owner.get("/manage/categories/").content.decode()
    assert _使用数(page, "NEWS") == 2
    assert _使用数(page, "カレンダー") == 2
    assert _使用数(page, "ショップ") == 1
    assert _使用数(page, "ホーム") == 1


def test_通知用カテゴリの割り当ては別名で当て_余りは順に配る():
    一覧 = [{"name": "にゅーす", "type": "通知"}, {"name": "お店", "type": "通知"},
          {"name": "Shop", "type": "通知"}, {"name": "教室", "type": "メニュー"}]
    割 = _通知カテゴリの割り当て(一覧)
    # 「にゅーす」はひらがな→カタカナ・小文字で「ニュース」の別名に当たる。Shop は shop。
    assert 割 == {"blog": "にゅーす", "product": "Shop", "calendar": "お店", "menu": "ホーム"}


# ── 追加・編集・削除 ─────────────────────────────────────


def test_足すと行番号は最大プラス1(as_owner, cats):
    r = as_owner.post("/manage/categories/add/", {"name": "商品情報", "kind": "通知"})
    assert r.status_code == 302
    c = Category.objects.get(name="商品情報")
    assert c.sheet_row == 6 and c.kind == "通知"


def test_同じ名前は足せない_失敗の文言は旧管理アプリと同じ(as_owner, cats):
    r = as_owner.post("/manage/categories/add/", {"name": "イベント情報", "kind": "ブログ"}, follow=True)
    assert Category.objects.filter(name="イベント情報").count() == 1
    assert "追加に失敗しました: 同じカテゴリ名が既に登録されています" in r.content.decode()


def test_名前と分類を変えられる_記事は変えない(as_owner, cats):
    News.objects.create(sheet_row=5, title="a", category="イベント情報", published=True)
    r = as_owner.post(
        "/manage/categories/update/",
        {"old_name": "イベント情報", "name": "イベント", "kind": "メニュー"}, follow=True,
    )
    assert Category.objects.filter(name="イベント", kind="メニュー").exists()
    assert News.objects.get(sheet_row=5).category == "イベント情報"
    assert "カテゴリを更新しました" in r.content.decode()


def test_行内の分類は種別お知らせならNEWSが選ばれている(as_owner, cats):
    page = as_owner.get("/manage/categories/").content.decode()
    row = _表の中身(page, "📰 NEWS用カテゴリ")
    first = row.split('value="イベント情報"')[1].split("</select>")[0]
    assert '<option value="ブログ" selected>NEWS</option>' in first


def test_使っていても消せる_記事の名前は残る(as_owner, cats):
    News.objects.create(sheet_row=5, title="a", category="イベント情報", published=True)
    r = as_owner.post("/manage/categories/delete/", {"name": "イベント情報"}, follow=True)
    assert not Category.objects.filter(name="イベント情報").exists()
    assert News.objects.get(sheet_row=5).category == "イベント情報"
    assert "カテゴリを削除しました" in r.content.decode()


def test_削除の前に旧管理アプリと同じ断り書きを出す(as_owner, cats):
    page = as_owner.get("/manage/categories/").content.decode()
    assert "既存データのカテゴリ名はそのまま残りますが、選択肢から外れます。" in page


def test_書き込みはGETでは動かない(as_owner, cats):
    assert as_owner.get("/manage/categories/delete/").status_code == 405
    assert as_owner.get("/manage/categories/add/").status_code == 405
    assert as_owner.get("/manage/categories/update/").status_code == 405
    assert Category.objects.count() == 4
