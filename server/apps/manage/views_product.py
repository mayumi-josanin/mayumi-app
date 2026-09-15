"""商品管理（ショップ）。**中身は旧管理アプリ（#page-products 商品マスター管理）と同じ。**

読み書きは gasapi/admin_product.py（GAS の転送先と同じ）。
原価（仕入値）は仕入の表（SupplierPrice）に商品名で持つ（GAS と同じ。商品の表には原価の列が無い）。
"""

from decimal import Decimal

from django.contrib import messages
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Category, Product
from apps.gasapi import admin_product
from apps.gasapi.views import _画像
from apps.records.models import SupplierPrice

from . import images, push
from .permissions import owner_required

SOLD_OUT = ["在庫あり", "売切"]


def _候補の分類():
    名 = set(Category.objects.exclude(kind__in=["お知らせ", "ブログ", "通知", "メニュー"]).values_list("name", flat=True))
    名 |= set(Product.objects.filter(deleted=False).exclude(category="").values_list("category", flat=True))
    return sorted(n for n in 名 if n)


def _原価を書く(名: str, 値: str):
    """仕入値を仕入の表へ（商品名で）。空なら触らない。GAS の handleAddProduct/UpdateProduct と同じ。"""
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


def _売切か(p: Product) -> bool:
    return (p.sold_out or "").strip() == "売切"


def _一件(p: Product, 原価: dict) -> dict:
    画像 = _画像(p.icon_url)
    stock = 0 if p.stock is None else p.stock
    warn = p.stock_warning or 0
    return {
        "p": p, "image": 画像[0] if 画像 else "", "icon": "" if 画像 else (p.icon_url or "🌿"),
        "cost": 原価.get((p.name or "").strip(), 0), "sold_out": _売切か(p), "stock": stock,
        "low": (not _売切か(p)) and warn > 0 and stock <= warn,
    }


@owner_required
def product_list(request):
    原価 = admin_product._原価の表()
    q = (request.GET.get("q") or "").strip()
    status = request.GET.get("status", "all")
    all_items = [_一件(p, 原価) for p in Product.objects.filter(deleted=False).order_by("sheet_row")]
    shown = []
    for it in all_items:
        p = it["p"]
        if status == "public" and not p.published:
            continue
        if status == "private" and p.published:
            continue
        if q and q not in " ".join([p.category or "", p.name or "", p.description or ""]):
            continue
        shown.append(it)
    summary = [
        ("商品総数", len(all_items)),
        ("公開中", sum(1 for it in all_items if it["p"].published)),
        ("非公開", sum(1 for it in all_items if not it["p"].published)),
        ("売切", sum(1 for it in all_items if it["sold_out"])),
        ("現在表示中", len(shown)),
    ]
    drafts = [it for it in shown if not it["p"].published]
    published = [it for it in shown if it["p"].published]
    any_drafts = any(not it["p"].published for it in all_items)
    any_published = any(it["p"].published for it in all_items)
    return render(request, "manage/product_list.html", {
        "drafts": drafts, "published": published, "summary": summary, "q": q, "status": status,
        "meta": (f"{len(shown)} / {len(all_items)} 件を表示" if all_items else "商品データはありません"),
        "drafts_empty": ("条件に一致する下書き商品はありません" if any_drafts else "下書き保存はありません"),
        "published_empty": ("条件に一致する公開商品はありません" if any_published else "投稿済み商品はありません"),
    })


def _入力(request, p: Product | None):
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "products") for f in request.FILES.getlist("images")]
    kept_desc = [u for u in request.POST.getlist("keep_desc_image") if u.strip()]
    added_desc = [images.保存する(f, "products") for f in request.FILES.getlist("desc_images")]
    status = request.POST.get("status") or "公開"
    return {
        "rowIdx": p.sheet_row if p else 0,
        "name": request.POST.get("name", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "price": request.POST.get("price", "0") or 0,
        "bg": request.POST.get("bg", "").strip(),
        "status": status,
        "description": request.POST.get("description", ""),
        "stockQty": request.POST.get("stockQty", "") or 0,
        "lowStockThreshold": request.POST.get("lowStockThreshold", "") or 0,
        "soldOutStatus": request.POST.get("soldOutStatus", "在庫あり"),
        "publishAt": request.POST.get("publishAt", ""),
        "imageUrls": kept + added,
        "icon": (kept + added)[0] if (kept + added) else request.POST.get("icon", "").strip(),
        "descriptionImageUrls": kept_desc + added_desc,
        "descriptionImage": (kept_desc + added_desc)[0] if (kept_desc + added_desc) else "",
    }


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


def _値(p: Product, 原価: dict) -> dict:
    return {
        "name": p.name, "category": p.category, "price": p.price or 0, "costPrice": 原価.get((p.name or "").strip(), ""),
        "bg": p.background_color or "#d4e8c8", "description": p.description,
        "stockQty": 0 if p.stock is None else p.stock, "lowStockThreshold": p.stock_warning or 0,
        "soldOutStatus": (p.sold_out or "").strip() or "在庫あり",
        "publishAt": timezone.localtime(p.publish_at).strftime("%Y-%m-%dT%H:%M") if p.publish_at else "",
        "icon": "" if _画像(p.icon_url) else (p.icon_url or ""),
    }


def _画面(request, p, values, existing, existing_desc, title, clone=False):
    return render(request, "manage/product_form.html", {
        "product": p, "values": values, "existing_images": existing, "existing_desc_images": existing_desc,
        "categories": _候補の分類(), "sold_out_choices": SOLD_OUT, "form_title": title, "clone": clone,
    })


@owner_required
def product_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        答 = admin_product.足す(d)
        if 答.get("status") == "ok":
            _原価を書く(d["name"], request.POST.get("costPrice", ""))
            messages.success(request, "商品を追加しました" if d["status"] == "公開" else "下書き保存しました")
            _通知(request, 答, d)
            return redirect("manage:product_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    return _画面(request, None, request.POST or {"soldOutStatus": "在庫あり", "bg": "#d4e8c8", "stockQty": 0, "lowStockThreshold": 0},
                [], [], "✨ 新しい商品を追加する")


@owner_required
def product_clone(request, row: int):
    """複製: 元の内容を「（複製）」付きの名前で新しい商品として入れる画面（cloneProduct）。"""
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    if request.method == "POST":
        return product_create(request)
    原価 = admin_product._原価の表()
    values = _値(p, 原価)
    values["name"] = f"{p.name}（複製）"
    return _画面(request, None, values, _画像(p.icon_url), _画像(p.description_image_url), "📄 商品を複製して追加する", clone=True)


@owner_required
def product_edit(request, row: int):
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, p)
        答 = admin_product.書き換える(d)
        if 答.get("status") == "ok":
            _原価を書く(d["name"], request.POST.get("costPrice", ""))
            messages.success(request, "商品を更新しました")
            _通知(request, 答, d)
            return redirect("manage:product_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    values = request.POST or _値(p, admin_product._原価の表())
    return _画面(request, p, values, _画像(p.icon_url), _画像(p.description_image_url), "📝 商品を編集する")


@owner_required
@require_POST
def product_row_save(request, row: int):
    """行内の「保存」: 価格・説明・公開設定・在庫・警告・売切だけ（saveProduct）。画像は触らない。"""
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    答 = admin_product.書き換える({
        "rowIdx": row, "price": request.POST.get("price", p.price or 0),
        "description": request.POST.get("description", p.description),
        "status": request.POST.get("status", "公開" if p.published else "非公開"),
        "stockQty": request.POST.get("stockQty", "") or 0,
        "lowStockThreshold": request.POST.get("lowStockThreshold", "") or 0,
        "soldOutStatus": request.POST.get("soldOutStatus", p.sold_out or "在庫あり"),
    })
    if 答.get("status") == "ok":
        messages.success(request, "商品情報を更新しました")
    else:
        messages.error(request, 答.get("message") or "更新できませんでした。")
    return redirect(request.POST.get("next") or "manage:product_list")


@owner_required
@require_POST
def product_delete(request, row: int):
    p = get_object_or_404(Product, sheet_row=row, deleted=False)
    答 = admin_product.消す({"rowIdx": row, "reason": f"管理画面から（{request.user.username}）"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{p.name}」をゴミ箱へ移動しました")
    return redirect("manage:product_list")


@owner_required
@require_POST
def product_bulk_delete(request):
    rows = [int(x) for x in request.POST.getlist("rows") if x.isdigit()]
    if not rows:
        messages.error(request, "削除する商品を選んでください。")
        return redirect(request.POST.get("next") or "manage:product_list")
    答 = admin_product.まとめて消す({"rowIdxs": rows, "reason": f"管理画面から一括（{request.user.username}）"})
    messages.success(request, f"{答.get('deleted', 0)}件をゴミ箱へ移動しました")
    return redirect(request.POST.get("next") or "manage:product_list")
