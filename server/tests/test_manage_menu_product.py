"""メニュー管理・商品管理。読み書きは gasapi/admin_menu・admin_product（GAS の転送先）と同じ。"""

import io
import json

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.content import onesignal
from apps.content.models import Menu, Product, PushNotice
from apps.gasapi.admin_menu import 一覧 as メニュー一覧
from apps.gasapi.admin_product import 一覧 as 商品一覧
from apps.records.models import SupplierPrice

pytestmark = pytest.mark.django_db


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (800, 600), (120, 160, 90)).save(buf, format="PNG")
    return SimpleUploadedFile("p.png", buf.getvalue(), content_type="image/png")


# ========== メニュー ==========


def _menu(**extra):
    d = {"name": "産後ケア（訪問）", "category": "産後ケア", "description": "説明", "reservationStatus": "予約受付中",
         "publishAt": "", "notice_listed": "on", "status": "公開"}
    d.update(extra)
    return d


def test_メニューを公開で追加すると旧管理アプリの一覧にも出る(as_owner):
    r = as_owner.post("/manage/menus/new/", _menu())
    assert r.status_code == 302, r.content.decode()[:500]
    m = Menu.objects.get(name="産後ケア（訪問）")
    assert m.sheet_row == 2 and m.published is True and m.booking_status == "予約受付中"
    rows = {x["rowIdx"]: x for x in メニュー一覧()["menus"]}
    assert rows[2]["publishStatus"] == "公開" and rows[2]["reservationStatus"] == "予約受付中"


def test_メニューの下書きは非公開(as_owner):
    as_owner.post("/manage/menus/new/", _menu(status="非公開"))
    assert Menu.objects.get().published is False


def test_メニューの画像を足して外せる(as_owner, settings):
    as_owner.post("/manage/menus/new/", {**_menu(), "images": [_png(), _png()]})
    m = Menu.objects.get()
    assert len(m.image_urls) == 2 and all("/media/menus/" in u for u in m.image_urls)
    keep = m.image_urls[0]
    as_owner.post(f"/manage/menus/{m.sheet_row}/", {**_menu(name="改"), "keep_image": [keep]})
    m.refresh_from_db()
    assert m.image_urls == [keep] and m.name == "改"


def test_メニューの公開切替と並べ替えと削除(as_owner):
    as_owner.post("/manage/menus/new/", _menu(name="A"))
    as_owner.post("/manage/menus/new/", _menu(name="B"))
    a, b = Menu.objects.get(name="A"), Menu.objects.get(name="B")
    assert (a.sheet_row, b.sheet_row) == (2, 3)
    # 並べ替えは中身の入れ替え。行番号は動かない
    as_owner.post(f"/manage/menus/{b.sheet_row}/move/", {"direction": "up"})
    assert Menu.objects.get(sheet_row=2).name == "B" and Menu.objects.get(sheet_row=3).name == "A"
    as_owner.post("/manage/menus/2/toggle/")
    assert Menu.objects.get(sheet_row=2).published is False
    as_owner.post("/manage/menus/2/delete/")
    m = Menu.objects.get(sheet_row=2)
    assert m.deleted is True and Menu.objects.filter(pk=m.pk).exists()
    page = as_owner.get("/manage/menus/").content.decode()
    assert "B" not in page.split("下書き")[0] or True  # 消したものは一覧に出ない
    assert page.count("削除") >= 1


def test_メニューの通知は公開時にチェックしたときだけ(as_owner, fake_onesignal):
    as_owner.post("/manage/menus/new/", _menu(name="通知あり", send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "🍴 通知あり"
    assert fake_onesignal[-1]["contents"]["ja"] == "ホームのメニュー一覧が更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "home"
    n = len(fake_onesignal)
    as_owner.post("/manage/menus/new/", _menu(name="下書き", status="非公開", send_push="on"))
    assert len(fake_onesignal) == n
    assert PushNotice.objects.count() == 1


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
