"""注文管理。**中身は旧管理アプリ（#page-orders）と同じ**（集計カード・絞り込み・一括更新・CSV・行の更新）。"""

import json

import pytest
from django.utils import timezone

from apps.content.models import Product
from apps.gasapi import orders
from apps.manage import analytics_calc as calc
from apps.records.models import OrderLine, SupplierPrice

pytestmark = pytest.mark.django_db


@pytest.fixture
def products(db):
    Product.objects.create(sheet_row=2, name="よもぎ茶（30パック）", price=1575, published=True)
    Product.objects.create(sheet_row=3, name="布良のクロス", price=2200, published=True)
    Product.objects.create(sheet_row=4, name="天然だし調味粉", price=2980, published=True)
    Product.objects.create(sheet_row=5, name="非公開の品", price=100, published=False)
    SupplierPrice.objects.create(sheet_row=2, product_name="よもぎ茶（30パック）", price=700)


@pytest.fixture
def order(products):
    答 = orders.注文する({"orderId": "ORD-1", "customerName": "山田 花子", "memberId": "MYM-0001",
                         "items": [{"name": "よもぎ茶（30パック）", "qty": 2, "price": 1575}], "total": 3150,
                         "payment": "現地払い"})
    assert 答["status"] == "ok"
    return "ORD-1"


# ========== サーバーの窓口（GAS も同じ経路） ==========


def test_商品を変えると行を作り直す(order):
    答 = orders.注文を書き換える({
        "orderId": order, "customerName": "山田 花子", "memberId": "MYM-0001",
        "items": [{"name": "よもぎ茶（30パック）", "qty": 1, "price": 1575}, {"name": "布良のクロス", "qty": 1, "price": 2200}],
        "total": 3775, "status": "受付中", "internalNote": "取り置き",
    })
    assert 答["status"] == "ok"
    行 = list(OrderLine.objects.filter(order_id=order).order_by("id"))
    assert [(r.product_name, r.quantity) for r in 行] == [("よもぎ茶（30パック）", 1), ("布良のクロス", 1)]
    assert 行[0].total_label == "¥3,775" and 行[0].internal_note == "取り置き" and 行[0].payment == "手動修正"
    assert int(行[0].cost_price) == 700 and int(行[0].profit) == 875


def test_手で入れた注文は日時と状態を持てる(products):
    答 = orders.注文する({"manual": True, "payment": "手動入力", "date": "2026-09-10T14:00",
                         "customerName": "佐藤", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}],
                         "total": 2200, "status": "受取済", "checked": True})
    r = OrderLine.objects.get(order_id=答["orderId"])
    assert r.payment == "手動入力" and r.status == "受取済" and r.received is True
    assert timezone.localtime(r.ordered_at).strftime("%Y-%m-%d %H:%M") == "2026-09-10 14:00"


# ========== 管理画面 ==========


def test_集計カードと既定の絞り込み(as_owner, order, products):
    orders.注文する({"orderId": "ORD-2", "customerName": "受取済の人", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}], "total": 2200})
    orders.注文を書き換える({"orderId": "ORD-2", "checked": True, "status": "受取済"})
    orders.注文する({"orderId": "ORD-3", "customerName": "取消の人", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}], "total": 2200})
    orders.取り消す({"orderId": "ORD-3"})
    page = as_owner.get("/manage/orders/").content.decode()
    for label in ["注文総数", "受付中", "受取済", "現在表示中"]:
        assert label in page
    assert "山田 花子" in page and "受取済の人" not in page and "取消の人" not in page
    assert "1 / 3 件を表示" in page
    assert "ここに表示される注文金額はデータ分析に反映されません" in page
    page = as_owner.get("/manage/orders/?showAll=1").content.decode()
    assert "受取済の人" in page and "取消の人" not in page   # キャンセル済は常に隠す
    page = as_owner.get("/manage/orders/?showAll=1&status=received").content.decode()
    assert "受取済の人" in page and "山田 花子" not in page
    page = as_owner.get("/manage/orders/?q=MYM-0001").content.decode()
    assert "山田 花子" in page
    page = as_owner.get("/manage/orders/?product=布良のクロス").content.decode()
    assert "山田 花子" not in page


def test_行の更新は状態と受取確認(as_owner, order):
    as_owner.post(f"/manage/orders/{order}/status/", {"status": "受付中", "checked": ""})
    assert not OrderLine.objects.get(order_id=order).received
    r = as_owner.post(f"/manage/orders/{order}/status/", {"status": "受取済", "checked": "1"}, follow=True)
    assert all(x.received for x in OrderLine.objects.filter(order_id=order))
    assert "受取確認済み — 一覧から削除しました" in r.content.decode()


def test_一括ステータス更新(as_owner, products):
    for i in range(3):
        orders.注文する({"orderId": f"B-{i}", "customerName": "x", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}], "total": 2200})
    as_owner.post("/manage/orders/bulk-status/", {"bulkStatus": "受取済", "selected": ["B-0", "B-2"]})
    assert OrderLine.objects.get(order_id="B-0").received and not OrderLine.objects.get(order_id="B-1").received
    assert OrderLine.objects.get(order_id="B-2").status == "受取済"


def test_CSVは全注文8列(as_owner, order):
    orders.注文を書き換える({"orderId": order, "internalNote": "メモ", "checked": True})
    r = as_owner.get("/manage/orders/csv/")
    assert r["Content-Type"].startswith("text/csv")
    body = r.content.decode("utf-8-sig")
    assert body.splitlines()[0] == "注文ID,日時,会員ID,氏名,商品,ステータス,受取確認,管理メモ"
    assert "ORD-1" in body and "よもぎ茶（30パック） x2" in body and ",済,メモ" in body


def test_新規注文の作成はだしの実質単価で合計(as_owner, products):
    r = as_owner.post("/manage/orders/new/", {
        "date": "2026-09-15", "customerName": "鈴木", "memberId": "", "status": "受付中", "internalNote": "電話注文",
        "items_json": json.dumps([{"name": "天然だし調味粉", "qty": 3, "price": 2980}]),
    })
    assert r.status_code == 302, r.content.decode()[:500]
    row = OrderLine.objects.get()
    assert row.customer_name == "鈴木" and row.quantity == 3 and row.payment == "手動入力"
    assert row.total_label == "¥" + f"{calc.dashi_pricing(3)['totalRevenue']:,}"
    page = as_owner.get("/manage/orders/new/").content.decode()
    assert "非公開の品" not in page and "布良のクロス (¥2200)" in page   # 公開商品だけ選べる


def test_商品が無ければ入れられない(as_owner, products):
    r = as_owner.post("/manage/orders/new/", {"customerName": "鈴木", "date": "2026-09-15", "items_json": "[]"})
    assert r.status_code == 200 and OrderLine.objects.count() == 0


def test_編集画面で商品を入れ替える(as_owner, order):
    page = as_owner.get(f"/manage/orders/{order}/edit/").content.decode()
    assert "📦 注文の編集" in page and 'value="山田 花子"' in page
    r = as_owner.post(f"/manage/orders/{order}/edit/", {
        "date": "2026-09-15", "customerName": "山田 花子", "memberId": "MYM-0001", "status": "受付中",
        "internalNote": "", "items_json": json.dumps([{"name": "布良のクロス", "qty": 2, "price": 2200}]),
    })
    assert r.status_code == 302
    行 = list(OrderLine.objects.filter(order_id=order))
    assert len(行) == 1 and 行[0].product_name == "布良のクロス" and 行[0].quantity == 2


def test_削除は本当に消す(as_owner, order):
    as_owner.post(f"/manage/orders/{order}/delete/")
    assert OrderLine.objects.filter(order_id=order).count() == 0
