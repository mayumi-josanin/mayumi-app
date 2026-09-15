"""商品管理。**中身は旧管理アプリ（商品マスター管理）と同じ**（集計カード5つ・検索と状態・行内編集・複製・一括削除）。"""

import io

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.content.models import Product
from apps.gasapi.admin_product import 一覧 as 商品一覧
from apps.records.models import SupplierPrice

pytestmark = pytest.mark.django_db


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (800, 600), (120, 160, 90)).save(buf, format="PNG")
    return SimpleUploadedFile("p.png", buf.getvalue(), content_type="image/png")


def _product(**extra):
    d = {"name": "よもぎ茶（30パック）", "category": "よもぎ茶", "price": "1575", "costPrice": "700",
         "bg": "#d4e8c8", "description": "説明", "stockQty": "12", "lowStockThreshold": "3",
         "soldOutStatus": "在庫あり", "publishAt": "", "icon": "🍵", "status": "公開"}
    d.update(extra)
    return d


def test_追加すると仕入値は仕入の表に商品名で入る(as_owner):
    r = as_owner.post("/manage/products/new/", _product())
    assert r.status_code == 302, r.content.decode()[:500]
    p = Product.objects.get(name="よもぎ茶（30パック）")
    assert p.price == 1575 and p.stock == 12 and p.stock_warning == 3 and p.background_color == "#d4e8c8" and p.icon_url == "🍵"
    assert int(SupplierPrice.objects.get(product_name="よもぎ茶（30パック）").price) == 700
    rows = {x["rowIdx"]: x for x in 商品一覧()["products"]}
    assert rows[p.sheet_row]["costPrice"] == 700 and rows[p.sheet_row]["bg"] == "#d4e8c8"


def test_集計カードと検索と状態の絞り込み(as_owner):
    as_owner.post("/manage/products/new/", _product())
    as_owner.post("/manage/products/new/", _product(name="下書きの品", status="非公開"))
    as_owner.post("/manage/products/new/", _product(name="売切の品", soldOutStatus="売切"))
    page = as_owner.get("/manage/products/").content.decode()
    for label in ["商品総数", "公開中", "非公開", "売切", "現在表示中", "商品マスター管理", "下書き保存一覧", "商品一覧 (価格変更・公開終了の操作)"]:
        assert label in page, label
    assert "SOLD OUT" in page and "アプリで購入不可" in page and "3 / 3 件を表示" in page
    page = as_owner.get("/manage/products/?status=private").content.decode()
    assert "下書きの品" in page and "売切の品" not in page and "条件に一致する公開商品はありません" in page
    page = as_owner.get("/manage/products/?q=売切").content.decode()
    assert "売切の品" in page and "下書きの品" not in page


def test_行内の保存は価格と説明と状態と在庫と売切(as_owner):
    as_owner.post("/manage/products/new/", _product())
    p = Product.objects.get()
    r = as_owner.post(f"/manage/products/{p.sheet_row}/row-save/", {
        "price": "1600", "description": "<strong>新</strong>説明", "status": "非公開", "stockQty": "2", "lowStockThreshold": "3", "soldOutStatus": "売切",
    })
    assert r.status_code == 302
    p.refresh_from_db()
    assert p.price == 1600 and p.description == "<strong>新</strong>説明" and p.published is False
    assert p.stock == 2 and p.sold_out == "売切" and p.icon_url == "🍵"  # 画像は触らない
    page = as_owner.get("/manage/products/").content.decode()
    assert "残り少なめ / 警告 3" not in page  # 売切のときは SOLD OUT 表示が優先


def test_在庫が閾値以下なら残り少なめ(as_owner):
    as_owner.post("/manage/products/new/", _product(stockQty="2", lowStockThreshold="3"))
    page = as_owner.get("/manage/products/").content.decode()
    assert "残り少なめ / 警告 3" in page and "在庫 2" in page


def test_複製は複製の名前で新しく追加(as_owner):
    as_owner.post("/manage/products/new/", {**_product(), "images": [_png()]})
    p = Product.objects.get()
    page = as_owner.get(f"/manage/products/{p.sheet_row}/clone/").content.decode()
    assert "📄 商品を複製して追加する" in page and 'value="よもぎ茶（30パック）（複製）"' in page and "/media/products/" in page
    r = as_owner.post(f"/manage/products/{p.sheet_row}/clone/", {**_product(name="よもぎ茶（30パック）（複製）"), "keep_image": [p.icon_url.split("\n")[0]]})
    assert r.status_code == 302
    assert Product.objects.count() == 2
    c = Product.objects.get(name="よもぎ茶（30パック）（複製）")
    assert c.sheet_row == p.sheet_row + 1 and "/media/products/" in c.icon_url


def test_編集画面のタイトルと仕入値(as_owner):
    as_owner.post("/manage/products/new/", _product())
    p = Product.objects.get()
    page = as_owner.get(f"/manage/products/{p.sheet_row}/").content.decode()
    assert "📝 商品を編集する" in page and 'value="700"' in page and "この保存でPush通知を送信する" in page


def test_一括削除は印だけ(as_owner):
    for n in ("a", "b", "c"):
        as_owner.post("/manage/products/new/", _product(name=n))
    rows = list(Product.objects.order_by("sheet_row").values_list("sheet_row", flat=True))
    as_owner.post("/manage/products/bulk-delete/", {"rows": [str(rows[0]), str(rows[2])]})
    assert Product.objects.filter(deleted=False).count() == 1 and Product.objects.count() == 3


def test_通知の文言(as_owner, fake_onesignal):
    as_owner.post("/manage/products/new/", _product(send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "🛍 よもぎ茶（30パック）"
    assert fake_onesignal[-1]["contents"]["ja"] == "ショップの商品情報が更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "shop"
