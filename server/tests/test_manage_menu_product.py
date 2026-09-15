"""商品管理（画像・原価）。読み書きは gasapi/admin_product（GAS の転送先）と同じ。メニューのテストは test_manage_menu.py。"""

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


# ========== 商品 ==========


def _product(**extra):
    d = {"name": "よもぎ茶（30パック）", "category": "よもぎ茶", "price": "1575", "costPrice": "700",
         "bg": "#d4e8c8", "description": "説明", "stockQty": "12", "lowStockThreshold": "3",
         "soldOutStatus": "在庫あり", "notice_listed": "on", "publishAt": "", "icon": "🍵", "status": "公開"}
    d.update(extra)
    return d



def test_商品の画像と説明画像(as_owner):
    as_owner.post("/manage/products/new/", {**_product(), "images": [_png()], "desc_images": [_png(), _png()]})
    p = Product.objects.get()
    assert "/media/products/" in p.icon_url
    assert p.description_image_url.count("/media/products/") == 2
    # 修正で説明画像を1枚だけ残す
    keep = p.description_image_url.split("\n")[0]
    as_owner.post(f"/manage/products/{p.sheet_row}/", {**_product(), "keep_image": [p.icon_url], "keep_desc_image": [keep]})
    p.refresh_from_db()
    assert p.description_image_url == keep and "/media/products/" in p.icon_url




def test_商品の原価を空で送っても消えない(as_owner):
    as_owner.post("/manage/products/new/", _product())
    p = Product.objects.get()
    as_owner.post(f"/manage/products/{p.sheet_row}/", _product(costPrice=""))
    assert int(SupplierPrice.objects.get(product_name=p.name).price) == 700
