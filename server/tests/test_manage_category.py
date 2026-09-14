"""カテゴリ管理。旧管理アプリと同じ決まり（名前が鍵・同じ名前は足せない）。"""

import pytest

from apps.content.models import Category, News

pytestmark = pytest.mark.django_db


@pytest.fixture
def cats(db):
    Category.objects.create(sheet_row=2, name="イベント情報", kind="お知らせ")
    Category.objects.create(sheet_row=3, name="まゆみのつぶやき", kind="ブログ")


def test_一覧に使用数が出る(as_owner, cats):
    News.objects.create(sheet_row=5, title="a", category="イベント情報", published=True)
    page = as_owner.get("/manage/categories/").content.decode()
    assert "イベント情報" in page
    assert "1件" in page


def test_足すと行番号は最大プラス1(as_owner, cats):
    r = as_owner.post("/manage/categories/add/", {"name": "商品情報", "kind": "お知らせ"})
    assert r.status_code == 302
    c = Category.objects.get(name="商品情報")
    assert c.sheet_row == 4 and c.kind == "お知らせ"


def test_同じ名前は足せない(as_owner, cats):
    as_owner.post("/manage/categories/add/", {"name": "イベント情報", "kind": "ブログ"})
    assert Category.objects.filter(name="イベント情報").count() == 1


def test_名前と種別を変えられる_記事は変えない(as_owner, cats):
    News.objects.create(sheet_row=5, title="a", category="イベント情報", published=True)
    as_owner.post(
        "/manage/categories/update/",
        {"old_name": "イベント情報", "name": "イベント", "kind": "ブログ"},
    )
    assert Category.objects.filter(name="イベント", kind="ブログ").exists()
    assert News.objects.get(sheet_row=5).category == "イベント情報"


def test_消せる(as_owner, cats):
    as_owner.post("/manage/categories/delete/", {"name": "まゆみのつぶやき"})
    assert not Category.objects.filter(name="まゆみのつぶやき").exists()


def test_書き込みはGETでは動かない(as_owner, cats):
    assert as_owner.get("/manage/categories/delete/").status_code == 405
    assert Category.objects.count() == 2
