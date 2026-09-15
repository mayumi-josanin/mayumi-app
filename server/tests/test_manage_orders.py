"""注文管理。読み書きは gasapi/orders（GAS の転送先）と同じ。"""

import json

import pytest
from django.utils import timezone

from apps.content.models import Product
from apps.gasapi import orders
from apps.records.models import OrderLine, SupplierPrice

pytestmark = pytest.mark.django_db


@pytest.fixture
def products(db):
    Product.objects.create(sheet_row=2, name="よもぎ茶（30パック）", price=1575, published=True)
    Product.objects.create(sheet_row=3, name="布良のクロス", price=2200, published=True)
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


def test_商品を送らなければ状態だけ変わる(order):
    orders.注文を書き換える({"orderId": order, "status": "受取済", "checked": True})
    行 = OrderLine.objects.filter(order_id=order)
    assert all(r.received for r in 行) and 行.first().status == "受取済"


def test_手で入れた注文は日時と状態を持てる(products):
    答 = orders.注文する({"manual": True, "payment": "手動入力", "date": "2026-09-10T14:00",
                         "customerName": "佐藤", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}],
                         "total": 2200, "status": "受取済", "checked": True})
    r = OrderLine.objects.get(order_id=答["orderId"])
    assert r.payment == "手動入力" and r.status == "受取済" and r.received is True
    assert timezone.localtime(r.ordered_at).strftime("%Y-%m-%d %H:%M") == "2026-09-10 14:00"


# ========== 管理画面 ==========


def test_一覧は受付中が既定で受取済は隠す(as_owner, order, products):
    orders.注文する({"orderId": "ORD-2", "customerName": "受取済の人", "items": [{"name": "布良のクロス", "qty": 1, "price": 2200}], "total": 2200})
    orders.注文を書き換える({"orderId": "ORD-2", "checked": True, "status": "受取済"})
    page = as_owner.get("/manage/orders/").content.decode()
    assert "山田 花子" in page and "受取済の人" not in page
    page = as_owner.get("/manage/orders/?status=received").content.decode()
    assert "受取済の人" in page
    page = as_owner.get("/manage/orders/?q=MYM-0001").content.decode()
    assert "山田 花子" in page


def test_一覧から状態と受取確認とメモ(as_owner, order):
    as_owner.post(f"/manage/orders/{order}/status/", {"internalNote": "月曜に来院"})
    assert OrderLine.objects.filter(order_id=order).first().internal_note == "月曜に来院"
    as_owner.post(f"/manage/orders/{order}/status/", {"checked": "1"})
    assert all(r.received for r in OrderLine.objects.filter(order_id=order))
    as_owner.post(f"/manage/orders/{order}/status/", {"status": "キャンセル"})
    assert OrderLine.objects.filter(order_id=order).first().status == "キャンセル"


def test_手で注文を入れる(as_owner, products):
    r = as_owner.post("/manage/orders/new/", {
        "date": "2026-09-15T10:00", "customerName": "鈴木", "memberId": "", "status": "受付中",
        "internalNote": "電話注文",
        "items_json": json.dumps([{"name": "よもぎ茶（30パック）", "qty": 3, "price": 1575}]),
    })
    assert r.status_code == 302, r.content.decode()[:500]
    r = OrderLine.objects.get()
    assert r.customer_name == "鈴木" and r.quantity == 3 and r.total_label == "¥4,725" and r.payment == "手動入力"
    assert r.order_id.startswith("ORD-")


def test_商品が無ければ入れられない(as_owner, products):
    r = as_owner.post("/manage/orders/new/", {"customerName": "鈴木", "items_json": "[]"})
    assert r.status_code == 200 and OrderLine.objects.count() == 0


def test_修正画面で商品を入れ替える(as_owner, order):
    page = as_owner.get(f"/manage/orders/{order}/edit/").content.decode()
    assert "よもぎ茶（30パック）" in page and 'value="山田 花子"' in page
    r = as_owner.post(f"/manage/orders/{order}/edit/", {
        "date": "2026-09-15T10:00", "customerName": "山田 花子", "memberId": "MYM-0001", "status": "受付中",
        "internalNote": "", "items_json": json.dumps([{"name": "布良のクロス", "qty": 2, "price": 2200}]),
    })
    assert r.status_code == 302
    行 = list(OrderLine.objects.filter(order_id=order))
    assert len(行) == 1 and 行[0].product_name == "布良のクロス" and 行[0].quantity == 2


def test_削除は本当に消す(as_owner, order):
    as_owner.post(f"/manage/orders/{order}/delete/")
    assert OrderLine.objects.filter(order_id=order).count() == 0
