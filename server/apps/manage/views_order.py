"""注文管理（ショップの注文）。

読み書きは gasapi/orders.py（GAS の転送先と同じ）。
一覧は注文IDごとにまとめ、既定では受取済みを隠す（旧管理アプリと同じ）。
"""

import json

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Product
from apps.gasapi import orders
from apps.records.models import OrderLine

STATUSES = ["受付中", "受取済", "キャンセル"]


def _状態(s: str) -> str:
    v = (s or "").strip()
    if not v or v in ("未受取", "pending"):
        return "受付中"
    if "受取" in v or v == "done":
        return "受取済"
    if "キャンセル" in v or v == "cancelled":
        return "キャンセル"
    return v


def _まとめ(show_all: bool):
    行 = OrderLine.objects.all().order_by("-ordered_at", "order_id", "id")
    if not show_all:
        行 = [r for r in 行 if not r.received]
    表, 順 = {}, []
    for r in 行:
        k = (r.order_id or "").strip()
        if not k:
            continue
        if k not in 表:
            順.append(k)
            表[k] = {
                "orderId": k, "date": r.ordered_at, "customerName": r.customer_name or "",
                "memberId": r.member_id or "", "status": _状態(r.status), "received": bool(r.received),
                "note": r.internal_note or "", "total": r.total_label or "", "payment": r.payment or "",
                "items": [], "amount": 0,
            }
        表[k]["items"].append({"name": r.product_name or "", "qty": r.quantity or 0,
                              "price": orders._金(r.unit_price)})
        表[k]["amount"] += (r.quantity or 0) * float(r.unit_price or 0)
    return [表[k] for k in 順]


@login_required
def order_list(request):
    status = request.GET.get("status", "pending")
    q = (request.GET.get("q") or "").strip()
    product = (request.GET.get("product") or "").strip()
    show_all = status in ("all", "received", "cancelled")
    一覧 = _まとめ(show_all)
    if status == "pending":
        一覧 = [o for o in 一覧 if o["status"] == "受付中" and not o["received"]]
    elif status == "received":
        一覧 = [o for o in 一覧 if o["received"] or o["status"] == "受取済"]
    elif status == "cancelled":
        一覧 = [o for o in 一覧 if o["status"] == "キャンセル"]
    if product:
        一覧 = [o for o in 一覧 if any(i["name"] == product for i in o["items"])]
    if q:
        一覧 = [o for o in 一覧 if q in o["customerName"] or q in o["memberId"] or q in o["orderId"]
                or any(q in i["name"] for i in o["items"])]
    全部 = _まとめ(True)
    summary = {
        "pending": sum(1 for o in 全部 if o["status"] == "受付中" and not o["received"]),
        "received": sum(1 for o in 全部 if o["received"] or o["status"] == "受取済"),
        "cancelled": sum(1 for o in 全部 if o["status"] == "キャンセル"),
    }
    products = sorted({i["name"] for o in 全部 for i in o["items"] if i["name"]})
    return render(request, "manage/order_list.html", {
        "orders": 一覧, "status": status, "q": q, "product": product, "products": products,
        "summary": summary, "statuses": STATUSES,
    })


@login_required
@require_POST
def order_status(request, order_id: str):
    """一覧からの1押し（状態・受取確認・メモ）。"""
    d = {"orderId": order_id}
    if "status" in request.POST:
        d["status"] = _状態(request.POST.get("status"))
        d["checked"] = d["status"] == "受取済"
    if "checked" in request.POST:
        d["checked"] = request.POST.get("checked") == "1"
    if "internalNote" in request.POST:
        d["internalNote"] = request.POST.get("internalNote", "")
    答 = orders.注文を書き換える(d)
    if 答.get("status") == "ok":
        messages.success(request, f"注文 {order_id} を更新しました。")
    else:
        messages.error(request, 答.get("message") or "更新できませんでした。")
    return redirect(request.POST.get("next") or "manage:order_list")


def _商品候補():
    return [(p.name, p.price or 0) for p in Product.objects.filter(deleted=False).order_by("sheet_row")]


def _品を読む(request) -> list:
    """items は JSON（画面の JS が組む）。"""
    try:
        品 = json.loads(request.POST.get("items_json") or "[]")
    except ValueError:
        品 = []
    出 = []
    for x in 品 if isinstance(品, list) else []:
        if isinstance(x, dict) and str(x.get("name") or "").strip():
            出.append({"name": str(x["name"]).strip(), "qty": int(x.get("qty") or 1), "price": int(float(x.get("price") or 0))})
    return 出


@login_required
def order_create(request):
    if request.method == "POST":
        品 = _品を読む(request)
        if not 品:
            messages.error(request, "商品を1つ以上入れてください。")
        else:
            答 = orders.注文する({
                "manual": True, "payment": "手動入力",
                "date": request.POST.get("date", ""), "customerName": request.POST.get("customerName", ""),
                "memberId": request.POST.get("memberId", ""), "items": 品,
                "total": sum(i["qty"] * i["price"] for i in 品),
                "status": _状態(request.POST.get("status")), "checked": request.POST.get("checked") == "1",
                "internalNote": request.POST.get("internalNote", ""),
            })
            if 答.get("status") == "ok":
                messages.success(request, f"注文 {答.get('orderId')} を作りました。")
                return redirect("manage:order_list")
            messages.error(request, 答.get("message") or "作れませんでした。")
    values = request.POST or {"date": timezone.localtime().strftime("%Y-%m-%dT%H:%M"), "status": "受付中"}
    return render(request, "manage/order_form.html", {
        "order": None, "values": values, "items": _品を読む(request) if request.method == "POST" else [],
        "products": _商品候補(), "statuses": STATUSES,
    })


@login_required
def order_edit(request, order_id: str):
    元 = [o for o in _まとめ(True) if o["orderId"] == order_id]
    if not 元:
        messages.error(request, "その注文は見つかりませんでした。")
        return redirect("manage:order_list")
    o = 元[0]
    if request.method == "POST":
        品 = _品を読む(request)
        if not 品:
            messages.error(request, "商品を1つ以上入れてください。")
        else:
            答 = orders.注文を書き換える({
                "orderId": order_id, "date": request.POST.get("date", ""),
                "customerName": request.POST.get("customerName", ""), "memberId": request.POST.get("memberId", ""),
                "items": 品, "total": sum(i["qty"] * i["price"] for i in 品),
                "status": _状態(request.POST.get("status")), "checked": request.POST.get("checked") == "1",
                "internalNote": request.POST.get("internalNote", ""),
            })
            if 答.get("status") == "ok":
                messages.success(request, f"注文 {order_id} を更新しました。")
                return redirect("manage:order_list")
            messages.error(request, 答.get("message") or "更新できませんでした。")
    values = request.POST or {
        "date": timezone.localtime(o["date"]).strftime("%Y-%m-%dT%H:%M") if o["date"] else "",
        "customerName": o["customerName"], "memberId": o["memberId"], "status": o["status"],
        "checked": "1" if o["received"] else "", "internalNote": o["note"],
    }
    return render(request, "manage/order_form.html", {
        "order": o, "values": values, "items": _品を読む(request) if request.method == "POST" else o["items"],
        "products": _商品候補(), "statuses": STATUSES,
    })


@login_required
@require_POST
def order_delete(request, order_id: str):
    答 = orders.注文を消す({"orderIds": [order_id]})
    if 答.get("status") == "ok":
        messages.success(request, f"注文 {order_id} を削除しました。")
    else:
        messages.error(request, 答.get("message") or "削除できませんでした。")
    return redirect("manage:order_list")
