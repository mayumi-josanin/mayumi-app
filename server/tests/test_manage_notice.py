"""お知らせ管理（お客様アプリの「お知らせ一覧」の横断ビュー）。**中身は旧管理アプリ（#page-notice-management）と同じ。**

集計カード6つ・検索と配信元の絞り込み・配信日時の新しい順・公開/非公開の切り替え・一覧から削除。
"""

import datetime

import pytest
from django.utils import timezone

from apps.content.models import CalendarEvent, Category, Menu, News, Product

pytestmark = pytest.mark.django_db


def _t(y, m, d, h=10):
    return timezone.make_aware(datetime.datetime(y, m, d, h, 0))


def _news(row, title, **extra):
    d = dict(sheet_row=row, posted_on=datetime.date(2026, 5, 1), title=title, body="本文", published=True,
             notice_listed=True, updated_at=_t(2026, 5, 1))
    d.update(extra)
    return News.objects.create(**d)


def _event(row, title, **extra):
    d = dict(sheet_row=row, event_on=datetime.date(2026, 6, 1), title=title, detail="詳細", published=True,
             notice_listed=True, updated_at=_t(2026, 6, 1))
    d.update(extra)
    return CalendarEvent.objects.create(**d)


def _product(row, name, **extra):
    d = dict(sheet_row=row, name=name, description="説明", published=True, notice_listed=True, updated_at=_t(2026, 7, 1))
    d.update(extra)
    return Product.objects.create(**d)


def _menu(row, name, **extra):
    d = dict(sheet_row=row, registered_on=datetime.date(2026, 8, 1), name=name, summary="概要", published=True,
             notice_listed=True, updated_at=_t(2026, 8, 1))
    d.update(extra)
    return Menu.objects.create(**d)


@pytest.fixture
def four_sources():
    _news(2, "米味噌作り")
    _event(2, "ベビーマッサージ教室", category="教室")
    _product(2, "よもぎ茶")
    _menu(2, "母乳相談")


def test_4つの配信元が配信日時の新しい順に並び集計カードが出る(as_owner, four_sources):
    page = as_owner.get("/manage/notices/").content.decode()
    for label in ["配信総数", "NEWS", "カレンダー", "ショップ", "ホーム", "現在表示中", "お知らせ管理",
                  "配信日時", "カテゴリ", "配信元", "タイトル / 内容", "お知らせ一覧", "操作",
                  "元データを編集", "一覧から削除", "4 / 4 件を表示",
                  "※「一覧から削除」を押すと、お知らせ一覧だけでなく ショップの商品一覧・ホームのメニュー一覧・カレンダーからも消えます。"]:
        assert label in page, label
    # 新しい順: ホーム(8/1) → ショップ(7/1) → カレンダー(6/1) → NEWS(5/1)
    順 = [page.index(t) for t in ["母乳相談", "よもぎ茶", "ベビーマッサージ教室", "米味噌作り"]]
    assert 順 == sorted(順)
    assert "2026/08/01" in page and "2026/05/01" in page
    # 元データの編集画面へのリンク
    for url in ["/manage/news/2/", "/manage/calendar/2/", "/manage/products/2/", "/manage/menus/2/"]:
        assert url in page, url


def test_出ないもの_非公開_一覧から削除済み_休診と往診_2026年4月より前(as_owner):
    _news(2, "出る記事")
    _news(3, "非公開の記事", published=False)
    _news(4, "一覧から外した記事", notice_delisted_at=_t(2026, 5, 2))
    _news(5, "昔の記事", posted_on=datetime.date(2026, 3, 1), updated_at=_t(2026, 3, 1))
    _news(6, "消した記事", deleted=True)
    _event(2, "休診のお知らせ", category="休診")
    _event(3, "往診の日", category="往診")
    _event(4, "夏のお休み", category="")
    _event(5, "出る教室", category="イベント", detail="4月からお休みしていた教室を再開")
    page = as_owner.get("/manage/notices/").content.decode()
    assert "出る記事" in page and "出る教室" in page
    for hidden in ["非公開の記事", "一覧から外した記事", "昔の記事", "消した記事", "休診のお知らせ", "往診の日", "夏のお休み"]:
        assert hidden not in page, hidden
    assert "2 / 2 件を表示" in page


def test_検索と配信元の絞り込みと空のときの文言(as_owner, four_sources):
    page = as_owner.get("/manage/notices/?source=product").content.decode()
    assert "よもぎ茶" in page and "米味噌作り" not in page and "1 / 4 件を表示" in page
    page = as_owner.get("/manage/notices/?q=味噌").content.decode()
    assert "米味噌作り" in page and "よもぎ茶" not in page
    page = as_owner.get("/manage/notices/?q=無い言葉").content.decode()
    assert "条件に一致する配信内容はありません" in page and "0 / 4 件を表示" in page
    # 検索は空白を無視し、長音とハイフンを同じに扱う（旧管理アプリの normalizeFilterText）
    page = as_owner.get("/manage/notices/?q=ベビー マッサージ").content.decode()
    assert "ベビーマッサージ教室" in page and "1 / 4 件を表示" in page


def test_1件も無いときの文言(as_owner):
    page = as_owner.get("/manage/notices/").content.decode()
    assert "管理対象のお知らせはありません" in page and "配信中のお知らせはありません" in page


def test_カテゴリは通知の種別のカテゴリ名を当てる(as_owner, four_sources):
    Category.objects.create(sheet_row=2, name="ニュース", kind="通知")
    Category.objects.create(sheet_row=3, name="お店から", kind="通知")
    page = as_owner.get("/manage/notices/").content.decode()
    # 「ニュース」は NEWS の別名で当たる。余った「お店から」は当たらなかった最初の配信元（カレンダー）へ
    assert ">ニュース<" in page and ">お店から<" in page and ">ショップ<" in page and ">ホーム<" in page


def test_本文の要約_印と画像行を外して180字で切る(as_owner):
    _news(2, "長い記事", body="<strong>太字</strong>\n📷 https://example.com/a.png\n" + "あ" * 200)
    _news(3, "空の記事", body="")
    page = as_owner.get("/manage/notices/").content.decode()
    assert "太字 " + "あ" * 177 + "…" in page and "<strong>太字" not in page and "example.com" not in page
    assert "本文なし" in page


def test_お知らせ一覧の公開非公開を切り替える(as_owner, four_sources):
    r = as_owner.post("/manage/notices/visibility/", {"kind": "product", "row": 2, "status": "非公開", "next": "/manage/notices/"})
    assert r.status_code == 302 and r.url == "/manage/notices/"
    p = Product.objects.get(sheet_row=2)
    assert p.notice_listed is False and p.notice_delisted_at is None and p.published is True  # 元データの公開は変えない
    # 非公開にしても一覧には残る（「一覧から削除」とは違う）。あとで公開に戻せる。
    page = as_owner.get("/manage/notices/").content.decode()
    assert "よもぎ茶" in page and "公開状態を更新しました" in page and 'value="非公開" selected' in page
    r = as_owner.post("/manage/notices/visibility/", {"kind": "calendar", "row": 2, "status": "公開"})
    assert r.status_code == 302
    c = CalendarEvent.objects.get(sheet_row=2)
    assert c.notice_listed is True and c.notice_listed_at is not None


def test_一覧から削除すると元データは残り一覧に出なくなる(as_owner, four_sources):
    r = as_owner.post("/manage/notices/remove/", {"kind": "menu", "row": 2})
    assert r.status_code == 302
    m = Menu.objects.get(sheet_row=2)
    assert m.deleted is False and m.notice_listed is False and m.notice_delisted_at is not None
    page = as_owner.get("/manage/notices/").content.decode()
    assert "母乳相談" not in page and "お知らせ一覧から完全に削除しました" in page and "3 / 3 件を表示" in page
    r = as_owner.post("/manage/notices/remove/", {"kind": "blog", "row": 2})
    assert News.objects.get(sheet_row=2).notice_delisted_at is not None


def test_おかしな宛先は断る(as_owner, four_sources):
    r = as_owner.post("/manage/notices/remove/", {"kind": "faq", "row": 2}, follow=True)
    assert "削除対象が見つかりません" in r.content.decode()
    r = as_owner.post("/manage/notices/visibility/", {"kind": "blog", "row": 99, "status": "非公開"}, follow=True)
    assert "公開状態の更新に失敗しました" in r.content.decode()
    assert as_owner.get("/manage/notices/visibility/").status_code == 405


def test_まゆみ以外は入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/notices/").status_code == 403
