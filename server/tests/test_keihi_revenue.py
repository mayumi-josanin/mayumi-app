"""売上の記録を、日ごとにまとめて現金出納帳へ入れる（院長の依頼 2026-09-21）。

    「9/20 にメニュー合計100,000円・商品合計50,000円なら、9/20 の売上は150,000円」

いちばん間違えたくないのは**お金の額**と、**院長が手で直した行を黙って書き換えないこと**。
"""

import datetime

import pytest

from apps.keihi.models import Cashbook, CashbookEntry
from apps.records.models import RevenueRecord

pytestmark = pytest.mark.django_db

下見 = "/manage/keihi/revenue/?year=2026&month=9"
取り込み = "/manage/keihi/revenue/to-book/?year=2026&month=9"


def 売上(kind, 日, 数, 単価, 名前="ためし", 行番号=None):
    return RevenueRecord.objects.create(
        kind=kind,
        sheet_row=行番号 if 行番号 is not None else RevenueRecord.objects.filter(kind=kind).count() + 2,
        recorded_on=datetime.date(2026, 9, 日),
        name=名前, quantity=数, unit_price=単価, unit_cost=0,
    )


@pytest.fixture
def 九月の売上():
    # 9/20: メニュー 100,000（50,000 × 2件）／商品 50,000 → 合計 150,000
    売上(RevenueRecord.MENU, 20, 1, 60000, "母乳外来")
    売上(RevenueRecord.MENU, 20, 1, 40000, "産後ケア")
    売上(RevenueRecord.PRODUCT, 20, 10, 5000, "よもぎ茶（30パック）")
    # 9/21: 商品だけ 3,000
    売上(RevenueRecord.PRODUCT, 21, 2, 1500, "よもぎ蒸し")
    # 9/22: 0円（無料・サービス）→ 行を作らない
    売上(RevenueRecord.PRODUCT, 22, 1, 0, "おまけ")
    # 10月は混ぜない
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=99,
                                 recorded_on=datetime.date(2026, 10, 1),
                                 name="来月", quantity=1, unit_price=777, unit_cost=0)


def test_下見に日ごとの合計が出る(as_owner, 九月の売上):
    page = as_owner.get(下見).content.decode()
    assert "150000" in page.replace(",", "")   # 9/20 メニュー10万＋商品5万
    assert "3000" in page.replace(",", "")     # 9/21
    assert "777" not in page                   # 10月は混ざらない


def test_0円の日は入れない(as_owner, 九月の売上):
    as_owner.post(取り込み, {"days": ["20", "21", "22"], "on_conflict": "skip"})
    日たち = set(CashbookEntry.objects.exclude(income=None).values_list("day", flat=True))
    assert 日たち == {20, 21}


def test_取り込むと1日1行できる(as_owner, 九月の売上):
    r = as_owner.post(取り込み, {"days": ["20", "21"], "on_conflict": "skip"})
    assert r.status_code == 302
    book = Cashbook.objects.get(year=2026, month=9)
    行 = [e for e in book.entries.all() if e.income]
    assert len(行) == 2
    二十日 = next(e for e in 行 if e.day == 20)
    assert 二十日.income == 150000            # メニュー10万 ＋ 商品5万
    assert 二十日.payment is None
    assert 二十日.counter_account == "売上高"
    assert 二十日.description.startswith("売上")


def test_二度押しても増えない(as_owner, 九月の売上):
    as_owner.post(取り込み, {"days": ["20", "21"], "on_conflict": "skip"})
    as_owner.post(取り込み, {"days": ["20", "21"], "on_conflict": "skip"})
    assert len([e for e in CashbookEntry.objects.all() if e.income]) == 2


def test_金額が変わったら選べる(as_owner, 九月の売上):
    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "skip"})
    売上(RevenueRecord.PRODUCT, 20, 1, 20000, "追加の商品")   # 9/20 が 170,000 になる

    page = as_owner.get(下見).content.decode()
    assert "金額が違います" in page

    # そのままにする → 帳簿は変わらない
    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "skip"})
    assert CashbookEntry.objects.filter(day=20).exclude(income=None).first().income == 150000

    # 直す → 新しい額になる
    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "replace"})
    assert CashbookEntry.objects.filter(day=20).exclude(income=None).first().income == 170000


def test_手で直した行は黙って書き換えない(as_owner, 九月の売上):
    """摘要を院長が直した行は「売上」で始まらないので、取り込みの対象にしない。"""
    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "skip"})
    e = CashbookEntry.objects.filter(day=20).exclude(income=None).first()
    e.description = "9/20 窓口でお預かり"
    e.income = 999
    e.save()

    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "replace"})
    e.refresh_from_db()
    assert e.income == 999 and e.description == "9/20 窓口でお預かり"


def test_月をまたがない(as_owner, 九月の売上):
    as_owner.post(取り込み, {"days": ["20", "21"], "on_conflict": "skip"})
    assert not Cashbook.objects.filter(month=10).exists()


def test_取り込んだ行は手で直せる(as_owner, 九月の売上):
    as_owner.post(取り込み, {"days": ["20"], "on_conflict": "skip"})
    e = CashbookEntry.objects.filter(day=20).exclude(income=None).first()
    e.income = 123456
    e.save()
    assert CashbookEntry.objects.get(pk=e.pk).income == 123456


def test_出納帳に取り込みの入口がある(as_owner):
    page = as_owner.get("/manage/keihi/?year=2026&month=9").content.decode()
    assert "この月の売上を取り込む" in page


def test_スタッフは入れない(client, staff):
    client.force_login(staff)
    assert client.get(下見).status_code in (302, 403)


def test_まとめての下見に月ごとの合計が出る(as_owner, 九月の売上):
    """6か月ぶん80日を1つずつ入れるのは大変なので、月をまとめて選べるようにした。"""
    売上(RevenueRecord.MENU, 5, 1, 200000, "8月のぶん")   # 9月以外も混ぜる
    RevenueRecord.objects.filter(name="8月のぶん").update(recorded_on=datetime.date(2026, 8, 5))

    page = as_owner.get("/manage/keihi/revenue/all/").content.decode()
    assert "2026年8月" in page and "2026年9月" in page
    assert "200000" in page.replace(",", "")


def test_まとめて取り込むと全部の月に入る(as_owner, 九月の売上):
    売上(RevenueRecord.MENU, 5, 1, 200000, "8月のぶん")
    RevenueRecord.objects.filter(name="8月のぶん").update(recorded_on=datetime.date(2026, 8, 5))

    r = as_owner.post("/manage/keihi/revenue/all/to-book/",
                      {"months": ["2026-8", "2026-9"], "on_conflict": "skip"})
    assert r.status_code == 302
    八月 = Cashbook.objects.get(year=2026, month=8)
    九月 = Cashbook.objects.get(year=2026, month=9)
    assert 八月.entries.exclude(income=None).count() == 1
    assert 八月.entries.exclude(income=None).first().income == 200000
    assert 九月.entries.exclude(income=None).count() == 2      # 9/20 と 9/21


def test_まとめても二度入れない(as_owner, 九月の売上):
    as_owner.post("/manage/keihi/revenue/all/to-book/", {"months": ["2026-9"], "on_conflict": "skip"})
    as_owner.post("/manage/keihi/revenue/all/to-book/", {"months": ["2026-9"], "on_conflict": "skip"})
    assert CashbookEntry.objects.exclude(income=None).count() == 2


def test_選ばなかった月は入らない(as_owner, 九月の売上):
    売上(RevenueRecord.MENU, 5, 1, 200000, "8月のぶん")
    RevenueRecord.objects.filter(name="8月のぶん").update(recorded_on=datetime.date(2026, 8, 5))

    as_owner.post("/manage/keihi/revenue/all/to-book/", {"months": ["2026-9"], "on_conflict": "skip"})
    assert not Cashbook.objects.filter(year=2026, month=8).exists()
