"""商品管理（ショップ）。

読み書きは gasapi/admin_product.py（GAS の転送先と同じ）。
原価は仕入の表（SupplierPrice）に商品名で持つ（GAS と同じ。商品の表には原価の列が無い）。
"""

from decimal import Decimal

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Category, Product
from apps.gasapi import admin_product
from apps.gasapi.views import _画像
from apps.records.models import SupplierPrice

from . import images, push

SOLD_OUT = ["在庫あり", "売切"]


def _候補の分類():
    名 = set(Category.objects.exclude(kind__in=["お知らせ", "ブログ", "通知", "メニュー"]).values_list("name", flat=True))
    名 |= set(Product.objects.filter(deleted=False).exclude(category="").values_list("category", flat=True))
    return sorted(n for n in 名 if n)


def _原価を書く(名: str, 値: str):
    """原価を仕入の表へ（商品名で）。空なら触らない。GAS の handleAddProduct/UpdateProduct と同じ。"""
    名 = (名 or "").strip()
    値 = (値 or "").strip()
    if not 名 or 値 == "":
        return
    try:
        金額 = Decimal(値)
    except Exception:
        return
    with transaction.atomic():
        s = SupplierPrice.objects.select_for_update().filter(product_name=名).order_by("sheet_row").first()
        if s:
            s.price = 金額
            s.save(update_fields=["price"])
        else:
            最大 = SupplierPrice.objects.order_by("-sheet_row").values_list("sheet_row", flat=True).first() or 1
            SupplierPrice.objects.create(sheet_row=最大 + 1, product_name=名, price=金額, memo="管理画面から追加")


@login_required
def product_list(request):
    原価 = admin_product._原価の表()
    rows = Product.objects.filter(deleted=False).order_by("sheet_row")
    items = []
    for p in rows:
        画像 = _画像(p.icon_url)
        items.append({
            "p": p, "image": 画像[0] if 画像 else "", "icon": "" if 画像 else (p.icon_url or "🌿"),
            "cost": 原価.get((p.name or "").strip(), 0),
            "sold_out": (p.sold_out or "").strip() == "売切",
            "low": (p.sold_out or "").strip() != "売切" and (p.stock_warning or 0) > 0
                   and p.stock is not None and 0 < p.stock <= (p.stock_warning or 0),
        })
    return render(request, "manage/product_list.html", {
        "published": [i for i in items if i["p"].published],
        "drafts": [i for i in items if not i["p"].published],
    })


def _入力(request, p: Product | None):
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "products") for f in request.FILES.getlist("images")]
    kept_desc = [u for u in request.POST.getlist("keep_desc_image") if u.strip()]
    added_desc = [images.保存する(f, "products") for f in request.FILES.getlist("desc_images")]
    status = request.POST.get("status") or "公開"
    d = {
        "rowIdx": p.sheet_row if p else 0,
        "name": request.POST.get("name", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "price": request.POST.get("price", "0") or 0,
        "bg": request.POST.get("bg", "").strip(),
        "status": status,
        "description": request.POST.get("description", ""),
        "stockQty": request.POST.get("stockQty", ""),
        "lowStockThreshold": request.POST.get("lowStockThreshold", ""),
        "soldOutStatus": request.POST.get("soldOutStatus", "在庫あり"),
        "noticeStatus": "公開" if request.POST.get("notice_listed") else "非公開",
        "publishAt": request.POST.get("publishAt", ""),
        "imageUrls": kept + added,
        "icon": (kept + added)[0] if (kept + added) else request.POST.get("icon", "").strip(),
        "descriptionImageUrls": kept_desc + added_desc,
        "descriptionImage": (kept_desc + added_desc)[0] if (kept_desc + added_desc) else "",
    }
    return d


def _通知(request, 答, d):
    """GAS の _商品の通知を送る_ と同じ条件。"""
    if not request.POST.get("send_push"):
        return
    状態 = 答.get("effectiveStatus") or d.get("status") or "公開"
    if 状態 == "非公開":
        return
    公開日時 = 答.get("effectivePublishAt") or ""
    if 公開日時:
        from apps.gasapi.admin_news import _日時

        t = _日時(公開日時)
        if t and t > timezone.now():
            return
    if push.全員へ送る("🛍 " + (d.get("name") or "商品更新"), "ショップの商品情報が更新されました", page="shop"):
        messages.info(request, "お客様のアプリへ通知を送りました。")
    else:
        messages.error(request, "通知を送れませんでした（商品は保存されています）。")


def _画面(request, p, values, existing, existing_desc):
    return render(request, "manage/product_form.html", {
        "product": p, "values": values, "existing_images": existing, "existing_desc_images": existing_desc,
        "categories": _候補の分類(), "sold_out_choices": SOLD_OUT,
    })


@login_required
def product_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        答 = admin_product.足す(d)
        if 答.get("status") == "ok":
            _原価を書く(d["name"], request.POST.get("costPrice", ""))
            messages.success(request, "商品を追加しました。" if d["status"] == "公開" else "下書きとして保存しました。")
            _通知(request, 答, d)
            return redirect("manage:product_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    return _画面(request, None, request.POST or {"soldOutStatus": "在庫あり", "bg": "#d4e8c8", "notice_listed": "on"}, [], [])


@login_required
def product_edit(request, row: int):
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, p)
        答 = admin_product.書き換える(d)
        if 答.get("status") == "ok":
            _原価を書く(d["name"], request.POST.get("costPrice", ""))
            messages.success(request, "商品を更新しました。")
            _通知(request, 答, d)
            return redirect("manage:product_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    原価 = admin_product._原価の表().get((p.name or "").strip(), "")
    values = request.POST or {
        "name": p.name, "category": p.category, "price": p.price or 0, "costPrice": 原価 or "",
        "bg": p.background_color or "#d4e8c8", "description": p.description,
        "stockQty": "" if p.stock is None else p.stock,
        "lowStockThreshold": "" if p.stock_warning is None else p.stock_warning,
        "soldOutStatus": (p.sold_out or "").strip() or "在庫あり",
        "notice_listed": "on" if p.notice_listed else "",
        "publishAt": timezone.localtime(p.publish_at).strftime("%Y-%m-%dT%H:%M") if p.publish_at else "",
        "icon": "" if _画像(p.icon_url) else (p.icon_url or ""),
    }
    return _画面(request, p, values, _画像(p.icon_url), _画像(p.description_image_url))


@login_required
@require_POST
def product_toggle(request, row: int):
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    答 = admin_product.公開を変える({"rowIdx": row, "status": "非公開" if p.published else "公開"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{p.name}」を{'非公開' if p.published else '公開'}にしました。")
    return redirect("manage:product_list")


@login_required
@require_POST
def product_sold_out(request, row: int):
    """売切 ⇄ 在庫あり。在庫数とは別（GAS と同じく、この印だけで判定する）。"""
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    次 = "在庫あり" if (p.sold_out or "").strip() == "売切" else "売切"
    答 = admin_product.書き換える({"rowIdx": row, "soldOutStatus": 次})
    if 答.get("status") == "ok":
        messages.success(request, f"「{p.name}」を{次}にしました。")
    return redirect("manage:product_list")


@login_required
@require_POST
def product_delete(request, row: int):
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    答 = admin_product.消す({"rowIdx": row, "reason": f"管理画面から（{request.user.username}）"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{p.name}」を削除しました。")
    return redirect("manage:product_list")
