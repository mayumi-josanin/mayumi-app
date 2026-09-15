"""データ分析と売上の記録。集計は gasapi/analytics、記録は gasapi/revenue（GAS の転送先）と同じ。"""

import json

import pytest

from apps.gasapi import orders
from apps.records.models import RevenueRecord, SupplierPrice

pytestmark = pytest.mark.django_db


@pytest.fixture
def data(db):
    SupplierPrice.objects.create(sheet_row=2, product_name="よもぎ茶（30パック）", price=700)
    orders.注文する({"orderId": "ORD-1", "customerName": "山田", "items": [{"name": "よもぎ茶（30パック）", "qty": 2, "price": 1575}], "total": 3150})
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=2, recorded_on="2026-09-10", name="母乳外来", quantity=3, unit_price=5000, unit_cost=0)
    RevenueRecord.objects.create(kind=RevenueRecord.PRODUCT, sheet_row=2, recorded_on="2026-09-11", name="布良のクロス", quantity=1, unit_price=2200, unit_cost=900)


def test_分析の月ごとと内訳が出る(as_owner, data):
    page = as_owner.get("/manage/analytics/").content.decode()
    assert "2026年9月" in page
    assert "母乳外来" in page and "3件" in page
    assert "布良のクロス" in page
    # メニュー 15,000 ＋ 商品（注文 3,150 ＋ 記録 2,200）
    assert "¥20,350" in page


def test_月を選ぶと内訳が変わる(as_owner, data):
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=3, recorded_on="2026-08-10", name="ビジリス", quantity=1, unit_price=3000, unit_cost=0)
    page = as_owner.get("/manage/analytics/?month=2026-08").content.decode()
    assert "2026年8月の内訳" in page and "¥3,000 × 1" in page


def test_メニュー収益をまとめて保存(as_owner):
    r = as_owner.post("/manage/revenue/menu/save/", {
        "date": "2026-09-15",
        "rows_json": json.dumps([{"name": "母乳外来", "qty": 2, "unitPrice": 5000, "unitCost": 0, "note": ""},
                                 {"name": "教室", "qty": 5, "unitPrice": 1500, "unitCost": 300, "note": "ベビマ"}]),
    })
    assert r.status_code == 302
    rows = list(RevenueRecord.objects.filter(kind=RevenueRecord.MENU).order_by("sheet_row"))
    assert [(x.name, x.quantity, int(x.unit_price)) for x in rows] == [("母乳外来", 2, 5000), ("教室", 5, 1500)]
    assert rows[0].sheet_row == 2 and rows[1].sheet_row == 3
    page = as_owner.get("/manage/revenue/menu/?month=2026-09").content.decode()
    assert "ベビマ" in page and "¥17,500" in page  # 合計


def test_商品収益の修正と削除で行番号が詰まる(as_owner):
    as_owner.post("/manage/revenue/product/save/", {
        "date": "2026-09-15",
        "rows_json": json.dumps([{"name": "A", "qty": 1, "unitPrice": 100, "unitCost": 10},
                                 {"name": "B", "qty": 1, "unitPrice": 200, "unitCost": 20},
                                 {"name": "C", "qty": 1, "unitPrice": 300, "unitCost": 30}]),
    })
    as_owner.post("/manage/revenue/product/save/", {"rowIdx": "3", "date": "2026-09-16", "name": "B改", "qty": "2", "unitPrice": "250", "unitCost": "20", "note": "直した"})
    b = RevenueRecord.objects.get(kind=RevenueRecord.PRODUCT, sheet_row=3)
    assert b.name == "B改" and b.quantity == 2 and int(b.unit_price) == 250
    as_owner.post("/manage/revenue/product/2/delete/")
    names = list(RevenueRecord.objects.filter(kind=RevenueRecord.PRODUCT).order_by("sheet_row").values_list("sheet_row", "name"))
    assert names == [(2, "B改"), (3, "C")]


def test_記録が無ければ保存しない(as_owner):
    r = as_owner.post("/manage/revenue/menu/save/", {"date": "2026-09-15", "rows_json": "[]"})
    assert r.status_code == 302 and RevenueRecord.objects.count() == 0


def test_知らない種別は分析へ戻す(as_owner):
    assert as_owner.get("/manage/revenue/other/")["Location"].endswith("/manage/analytics/")
