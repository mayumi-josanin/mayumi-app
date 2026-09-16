"""データ分析と売上の記録。**旧管理アプリの計算と同じ**（注文は入れない・だしの割引段階・2026年4月以降）。"""

import json

import pytest

from apps.content.models import Product
from apps.gasapi import orders
from apps.manage import analytics_calc as calc
from apps.records.models import RevenueRecord, SupplierPrice

pytestmark = pytest.mark.django_db


@pytest.fixture
def data(db):
    Product.objects.create(sheet_row=2, name="よもぎ茶（30パック）", price=1575, published=True, sort_order=1)
    Product.objects.create(sheet_row=3, name="布良のクロス", price=2200, published=True, sort_order=2)
    SupplierPrice.objects.create(sheet_row=2, product_name="よもぎ茶（30パック）", price=700)
    # ショップの注文は集計に入れない（旧アプリの注記どおり）
    orders.注文する({"orderId": "ORD-1", "customerName": "山田", "items": [{"name": "よもぎ茶（30パック）", "qty": 2, "price": 1575}], "total": 3150})
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=2, recorded_on="2026-09-10", name="母乳外来", quantity=3, unit_price=5000, unit_cost=0)
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=3, recorded_on="2026-09-12", name="母乳外来", quantity=1, unit_price=3000, unit_cost=0)
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=4, recorded_on="2026-09-12", name="骨盤健診", quantity=2, unit_price=4000, unit_cost=500)
    RevenueRecord.objects.create(kind=RevenueRecord.PRODUCT, sheet_row=2, recorded_on="2026-09-11", name="布良のクロス", quantity=1, unit_price=2200, unit_cost=900)
    RevenueRecord.objects.create(kind=RevenueRecord.PRODUCT, sheet_row=3, recorded_on="2026-08-05", name="布良のクロス", quantity=2, unit_price=2200, unit_cost=900)
    RevenueRecord.objects.create(kind=RevenueRecord.MENU, sheet_row=5, recorded_on="2026-03-05", name="教室", quantity=1, unit_price=1000, unit_cost=0)  # 4月より前は出さない


# ========== 計算（旧アプリの写し） ==========


def test_だしの割引段階():
    assert calc.dashi_tiers(2) == {"20": 2, "25": 0, "30": 0, "35": 0}
    assert calc.dashi_tiers(3) == {"20": 2, "25": 1, "30": 0, "35": 0}
    assert calc.dashi_tiers(4) == {"20": 3, "25": 0, "30": 1, "35": 0}
    assert calc.dashi_tiers(5) == {"20": 4, "25": 0, "30": 0, "35": 1}
    assert calc.dashi_tiers(6) == {"20": 0, "25": 6, "30": 0, "35": 0}
    p = calc.dashi_pricing(3)
    assert p["totalRevenue"] == 2380 * 2 + 2235 and p["text"] == "2,380円(20%OFF) × 2個 / 2,235円(25%OFF) × 1個"
    assert calc.dashi_pricing(0)["avgUnitPrice"] == 2980


def test_月の計算と比較():
    assert calc.shift_month("2026-01", -1) == "2025-12" and calc.shift_month("2026-09", -12) == "2025-09"
    assert calc.month_visible("2026-04") and not calc.month_visible("2026-03") and calc.month_visible("2027-01")
    assert calc.comparison(120, 100)["text"] == "+20.0%" and calc.comparison(80, 100)["cls"] == "down"
    assert calc.comparison(10, 0)["text"] == "-"
    assert calc.ratio_label(30, 120) == "25.0%" and calc.ratio_label(1, 0) == "-"
    assert calc.axis_label(15000) == "1.5万" and calc.axis_label(20000) == "2万" and calc.axis_label(120) == "120"
    assert calc.chart_ceiling(12345) == 20000 and calc.chart_ceiling(4000) == 5000


def test_集計は注文を入れず記録だけ(data):
    res = calc.build()
    m = res["matrix"]["2026-09"]
    assert m["sales"] == 2200 and m["profit"] == 1300           # 注文の 3,150 は入らない
    assert m["menuRevenueTotal"] == 15000 + 3000 + 8000
    assert m["menuCount"]["母乳外来"] == 4 and m["menuCount"]["その他"] == 2
    assert set(m["menuBreakdown"]["母乳外来"]) == {"5000", "3000"}
    assert "骨盤健診" in m["menuDetails"] and m["menuDetails"]["骨盤健診"]["profit"] == 7000
    assert m["combinedRevenue"] == 26000 + 2200
    assert "2026-03" in res["matrix"] and "布良のクロス" in res["products"]


def test_だしは価格帯ごとに分かれる(db):
    Product.objects.create(sheet_row=2, name="天然だし調味粉", price=2980, published=True)
    RevenueRecord.objects.create(kind=RevenueRecord.PRODUCT, sheet_row=2, recorded_on="2026-09-01", name="天然だし調味粉",
                                 quantity=3, unit_price=calc.dashi_pricing(3)["avgUnitPrice"], unit_cost=1490)
    m = calc.build()["matrix"]["2026-09"]
    keys = sorted(m["productDetails"])
    assert keys == ["天然だし調味粉::¥2,235 (25%OFF)", "天然だし調味粉::¥2,380 (20%OFF)"]
    assert m["productDetails"]["天然だし調味粉::¥2,380 (20%OFF)"]["qty"] == 2


# ========== 画面 ==========


def test_分析画面に旧アプリと同じ区画が出る(as_owner, data):
    page = as_owner.get("/manage/analytics/").content.decode()
    for h in ["📈 月別比較", "📊 粗利サマリー", "📊 月別データグラフ", "💰 収益サマリー (月別)", "🩺 メニュー別収益 (月別)",
              "📦 商品別収益詳細 (月別)", "📝 商品別販売数 (推移)"]:
        assert h in page, h
    assert "2026-09 総売上" in page and "¥28,200" in page
    assert "2026-03" not in page                       # 4月より前は出ない
    assert page.count("<svg") == 4                     # 棒グラフ4本
    assert "全メニュー合計" in page and "骨盤健診" in page  # グラフの選択肢
    assert "+" in page or "-" in page                  # 前月比


def test_月を選ぶと商品とメニューの内訳(as_owner, data):
    page = as_owner.get("/manage/analytics/?month=2026-09").content.decode()
    assert "🍴 メニュー別収益詳細 (月別)" in page
    assert "布良のクロス" in page and "¥2,200" in page and "1個" in page
    assert "骨盤健診" in page and "¥4,000: 2件" in page


def test_グラフの対象を選べる(as_owner, data):
    page = as_owner.get("/manage/analytics/?menu_target=骨盤健診").content.decode()
    assert "「骨盤健診」の月ごとの売上" in page


# ========== 記録 ==========


def test_メニュー収益をまとめて保存とその他の名前(as_owner):
    r = as_owner.post("/manage/revenue/menu/save/", {
        "date": "2026-09-15",
        "rows_json": json.dumps([{"name": "母乳外来", "qty": 2, "unitPrice": 5000, "unitCost": 0, "note": ""},
                                 {"name": "骨盤健診", "qty": 5, "unitPrice": 1500, "unitCost": 300, "note": "ベビマ"}]),
    })
    assert r.status_code == 302
    rows = list(RevenueRecord.objects.filter(kind=RevenueRecord.MENU).order_by("sheet_row"))
    assert [(x.name, x.quantity) for x in rows] == [("母乳外来", 2), ("骨盤健診", 5)]
    page = as_owner.get("/manage/revenue/menu/?month=2026-09").content.decode()
    assert "ベビマ" in page and "2行 / 合計7件" in page and "登録済みメニュー収益" in page


def test_だしの商品収益は単価が自動(as_owner):
    Product.objects.create(sheet_row=2, name="天然だし調味粉", price=2980, published=True)
    as_owner.post("/manage/revenue/product/save/", {
        "date": "2026-09-15", "rows_json": json.dumps([{"name": "天然だし調味粉", "qty": 3, "unitPrice": 2980, "unitCost": 1490}]),
    })
    rec = RevenueRecord.objects.get()
    assert round(float(rec.unit_price)) == round(calc.dashi_pricing(3)["avgUnitPrice"])
    page = as_owner.get("/manage/revenue/product/").content.decode()
    assert "実質" in page and "2,380円(20%OFF) × 2個" in page


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


def test_メニュー別収益は月で絞れる_すべても選べる(as_owner, data):
    """院長の希望（2026-09-16）: メニュー別収益(月別)を月で検索。「すべて」で全部の月。"""
    page = as_owner.get("/manage/analytics/").content.decode()
    i = page.find("メニュー別収益 (月別)")
    assert '<option value="" selected>すべて' in page[i:]
    assert "2026/09" in page[i:] and "2026/08" in page[i:]
    page = as_owner.get("/manage/analytics/?menu_month=2026-09").content.decode()
    i = page.find("メニュー別収益 (月別)")
    j = page.find("商品別収益", i)
    区画 = page[i:j]
    assert "<td class=\"nowrap\">2026-09</td>" in 区画 and "<td class=\"nowrap\">2026-08</td>" not in 区画
    assert '<option value="2026-09" selected>' in 区画
    # 無い月を指定しても落ちず、すべてに戻る
    page = as_owner.get("/manage/analytics/?menu_month=1999-01").content.decode()
    assert '<option value="" selected>すべて' in page
