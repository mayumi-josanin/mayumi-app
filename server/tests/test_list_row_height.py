"""一覧の1行の縦幅を、管理画面の全部の一覧でそろえる（院長の希望 2026-09-19）。

これまでは画面ごとに言われるたびに1つずつそろえてきた（カレンダー・NEWS・お知らせ・
メニューは64px、会員は84px）。**画面ごとに書き方も高さも違っていた。**
いまは共通の決まり（`.list-table` と `.list-line`）を style.css に1か所だけ置き、
一覧の表はどの画面もそれを使う。ここで見張るのは次の3つ。

1. 共通の決まりが style.css に1つだけあること（同じことが2か所にない）
2. 一覧の画面を開くと、その表に共通の決まりが当たっていること
3. 中身が多い行でも、高さの指定は同じ（はみ出す文字は「…」で省く）こと
"""

from datetime import date, datetime
from pathlib import Path

import pytest
from django.utils import timezone

from apps.content.models import CalendarEvent, Menu, News, Product
from apps.dev.models import DevProject, DevTask
from apps.records.models import RevenueRecord
from apps.members.models import Member
from apps.content.models import PushNotice
from apps.gasapi import orders as 注文の窓口

pytestmark = pytest.mark.django_db

CSS = Path(__file__).resolve().parent.parent / "apps" / "manage" / "static" / "manage" / "style.css"

# 一覧が並ぶ画面。左が住所、右が画面の呼び名（うまくいかなかったときに分かるように）
一覧の画面 = [
    ("/manage/members/", "会員一覧"),
    ("/manage/news/", "NEWS"),
    ("/manage/notices/", "お知らせ"),
    ("/manage/menus/", "メニュー"),
    ("/manage/products/", "商品"),
    ("/manage/calendar/", "カレンダー"),
    ("/manage/orders/", "注文"),
    ("/manage/push/", "通知"),
    ("/manage/categories/", "カテゴリ"),
    ("/manage/trash/", "ゴミ箱"),
    ("/manage/rewards/", "スタンプ・特典"),
    ("/manage/revenue/menu/", "売上（メニュー）"),
    ("/manage/dev/tasks/", "開発のタスク"),
    ("/manage/dev/", "開発のプロジェクト"),
    ("/manage/bijiris/customers/", "ビジリスの顧客"),
    ("/manage/bijiris/rewards/", "ビジリスの特典"),
]


def _会員(member_id, name, **extra):
    d = {"kana": "", "phone": "", "address": "", "created_at": timezone.now()}
    d.update(extra)
    return Member.objects.create(member_id=member_id, name=name, **d)


def _t(y, m, d, h=10):
    return timezone.make_aware(datetime(y, m, d, h, 0))


@pytest.fixture
def 一覧の中身(db):
    """どの一覧にも1件ずつ置く。中身が空でも表は出るが、行のある状態で確かめたい。"""
    _会員("MYM-0001", "山田花子")
    News.objects.create(sheet_row=2, posted_on=date(2026, 5, 1), title="お知らせ", body="本文",
                        published=True, notice_listed=True, updated_at=_t(2026, 5, 1))
    CalendarEvent.objects.create(sheet_row=2, event_on=date(2026, 6, 1), title="教室", detail="詳細",
                                 published=True, notice_listed=True, updated_at=_t(2026, 6, 1))
    Menu.objects.create(sheet_row=2, name="産後ケア", summary="説明", published=True, updated_at=_t(2026, 6, 1))
    Product.objects.create(sheet_row=2, name="ハーブティー", description="説明", published=True, updated_at=_t(2026, 7, 1))
    PushNotice.objects.create(sheet_row=2, title="通知", body="本文", sent_at=_t(2026, 7, 1))
    注文の窓口.注文する({"orderId": "ORD-1", "customerName": "山田花子", "memberId": "MYM-0001",
                       "items": [{"name": "ハーブティー", "qty": 1, "price": 100}], "total": 100,
                       "payment": "現地払い"})
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=2, recorded_on="2026-09-10",
                                 name="母乳外来", quantity=1, unit_price=5000, unit_cost=0)
    p = DevProject.objects.create(name="お客様アプリ", status="in_progress")
    DevTask.objects.create(project=p, title="一覧の縦幅をそろえる", status="open", priority="high")


def test_共通の決まりはstyle_cssに1か所だけある():
    """画面ごとの高さの指定が残っていると、また画面ごとにずれていく。"""
    css = CSS.read_text(encoding="utf-8")
    assert css.count(".list-table td { height:") == 1
    assert ".list-table .list-line {" in css
    # 画面ごとにばらばらだった昔の指定は消したままにする
    for 昔の指定 in (".cal-table td {", ".news-table td {", ".notice-table td {",
                     ".menu-table td {", ".member-table td {"):
        assert 昔の指定 not in css


def test_高さは84pxで会員の氏名の3行が収まる():
    """会員一覧の氏名の欄が3行（お名前・注文/受付中/端末・在席と最終オンライン）で、
    ここが一覧の中でいちばん背が高い。84pxはその3行が収まる中でいちばん小さい値。
    """
    css = CSS.read_text(encoding="utf-8")
    assert ".list-table td { height: 84px;" in css


@pytest.mark.parametrize("url,画面", 一覧の画面)
def test_どの一覧にも共通の決まりが当たっている(as_owner, 一覧の中身, url, 画面):
    r = as_owner.get(url)
    assert r.status_code == 200, 画面
    assert 'class="list-table"' in r.content.decode(), 画面


def test_中身の多い行でも高さの指定は同じ(as_owner):
    """行の中身が多い会員も少ない会員も、同じ高さで並ぶこと。

    はみ出す文字は .list-line で「…」に省き、全文は title（ホバー）で読める。
    """
    住所 = "神奈川県厚木市中町1-1-1 コーポ山田303号室"
    _会員("MYM-0001", "山田花子", address=住所, memo="長い覚え書き" * 10)
    _会員("MYM-0002", "佐藤桃子")
    page = as_owner.get("/manage/members/").content.decode()
    assert 'class="list-table"' in page
    assert "list-line" in page
    assert f'title="{住所}"' in page


def test_注文の商品は2行までで残りは件数にまとめる(as_owner):
    """商品の数で行の高さが変わらないように、3件目からは「ほか N件」にする。"""
    Product.objects.create(sheet_row=2, name="ハーブティー", published=True, price=100)
    注文の窓口.注文する({"orderId": "ORD-1", "customerName": "山田花子", "memberId": "MYM-0001",
                       "items": [{"name": f"商品{i}", "qty": 1, "price": 100} for i in range(5)],
                       "total": 500, "payment": "現地払い"})
    page = as_owner.get("/manage/orders/").content.decode()
    assert "ほか 3件" in page
    # 3件目からは行に出さない（全部はホバーの title で読める）
    assert '<div class="list-line">商品2' not in page
