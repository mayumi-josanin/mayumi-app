"""注文管理（ショップの注文）。**中身は旧管理アプリ（#page-orders）と同じ。**

読み書きは gasapi/orders.py（GAS の転送先と同じ）。
一覧は注文IDごとにまとめ、キャンセル済は常に隠し、受取済は「受取済も含めて表示」のときだけ出す。
"""

import csv
import json

from django.contrib import messages
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Product
from apps.gasapi import orders
from apps.records.models import OrderLine

from . import analytics_calc as calc
from .permissions import owner_required

STATUSES = ["受付中", "受取済"]


def _状態(s: str) -> str:
    """normalizeAdminOrderStatus。"""
    v = (s or "").strip()
    if not v or v in ("未受取", "受付中", "pending"):
        return "受付中"
    if v == "受取済" or "受取済" in v or v == "done":
        return "受取済"
    if v == "キャンセル済" or "キャンセル" in v or v == "cancelled":
        return "キャンセル済"
    return v


def _まとめ():
    """全部の注文（getAdminOrders showAll=true 相当）。API の並び（古い順）のまま。"""
    行 = OrderLine.objects.all().order_by("id")
    表, 順 = {}, []
    for r in 行:
        k = (r.order_id or "").strip()
        if not k:
            continue
        if k not in 表:
            順.append(k)
            表[k] = {
                "orderId": k, "date": r.ordered_at, "customerName": r.customer_name or "",
                "memberId": r.member_id or "", "status": _状態(r.status), "checked": bool(r.received),
                "note": r.internal_note or "", "total": r.total_label or "", "items": [],
            }
        表[k]["items"].append({"name": r.product_name or "", "qty": r.quantity or 0, "price": orders._金(r.unit_price)})
    return [表[k] for k in 順]


def _絞る(all_orders, show_all, status, product, q):
    """applyOrderFilters。"""
    out = []
    for o in all_orders:
        if o["status"] == "キャンセル済":
            continue
        if not show_all and (o["checked"] or o["status"] == "受取済"):
            continue
        if status == "pending" and o["status"] != "受付中":
            continue
        if status == "received" and o["status"] != "受取済":
            continue
        if product and not any(i["name"] == product for i in o["items"]):
            continue
        if q:
            names = " ".join(i["name"] for i in o["items"])
            hay = " ".join([o["orderId"], o["memberId"], o["customerName"], names])
            if q not in hay:
                continue
        out.append(o)
    return out


@owner_required
def order_list(request):
    show_all = request.GET.get("showAll") == "1"
    status = request.GET.get("status", "all")
    product = (request.GET.get("product") or "").strip()
    q = (request.GET.get("q") or "").strip()
    all_orders = _まとめ()
    shown = _絞る(all_orders, show_all, status, product, q)
    summary = [
        ("注文総数", len(all_orders)),
        ("受付中", sum(1 for o in all_orders if o["status"] == "受付中")),
        ("受取済", sum(1 for o in all_orders if o["status"] == "受取済")),
        ("現在表示中", len(shown)),
    ]
    products = sorted({i["name"] for o in all_orders for i in o["items"] if i["name"]})
    meta = f"{len(shown)} / {len(all_orders)} 件を表示" if all_orders else "注文データはありません"
    return render(request, "manage/order_list.html", {
        "orders": shown, "all_count": len(all_orders), "show_all": show_all, "status": status,
        "product": product, "q": q, "products": products, "summary": summary, "meta": meta, "statuses": STATUSES,
    })


@owner_required
@require_POST
def order_status(request, order_id: str):
    """行の「更新」: ステータスの選択と受取確認をそのまま保存（saveOrder）。"""
    status = _状態(request.POST.get("status"))
    checked = request.POST.get("checked") == "1"
    答 = orders.注文を書き換える({"orderId": order_id, "status": status, "checked": checked})
    if 答.get("status") == "ok":
        messages.success(request, "受取確認済み — 一覧から削除しました" if checked else "注文情報を更新しました")
    else:
        messages.error(request, 答.get("message") or "更新できませんでした。")
    return redirect(request.POST.get("next") or "manage:order_list")


@owner_required
@require_POST
def order_bulk_status(request):
    """一括ステータス更新（bulkUpdateOrdersStatus）。"""
    status = _状態(request.POST.get("bulkStatus"))
    ids = [x for x in request.POST.getlist("selected") if x.strip()]
    if not ids or not request.POST.get("bulkStatus"):
        messages.error(request, "注文を選び、一括変更の状態を選んでください。")
        return redirect(request.POST.get("next") or "manage:order_list")
    for oid in ids:
        orders.注文を書き換える({"orderId": oid, "status": status, "checked": status == "受取済"})
    messages.success(request, "注文ステータスを一括更新しました")
    return redirect(request.POST.get("next") or "manage:order_list")


@owner_required
def order_csv(request):
    """CSV出力（exportOrdersCsv）: 全注文、8列。"""
    res = HttpResponse(content_type="text/csv; charset=utf-8-sig")
    res["Content-Disposition"] = 'attachment; filename="orders.csv"'
    w = csv.writer(res)
    w.writerow(["注文ID", "日時", "会員ID", "氏名", "商品", "ステータス", "受取確認", "管理メモ"])
    for o in _まとめ():
        w.writerow([o["orderId"], timezone.localtime(o["date"]).strftime("%Y/%m/%d %H:%M") if o["date"] else "",
                    o["memberId"], o["customerName"],
                    " / ".join(f"{i['name']} x{i['qty']}" for i in o["items"]),
                    o["status"], "済" if o["checked"] else "", o["note"]])
    return res


def _商品候補():
    return [(p.name, p.price or 0) for p in Product.objects.filter(deleted=False, published=True).order_by("sheet_row")]


def _品を読む(request) -> list:
    try:
        品 = json.loads(request.POST.get("items_json") or "[]")
    except ValueError:
        品 = []
    出 = []
    for x in 品 if isinstance(品, list) else []:
        if isinstance(x, dict) and str(x.get("name") or "").strip():
            name = str(x["name"]).strip()
            qty = int(x.get("qty") or 1)
            price = float(x.get("price") or 0)
            if calc.is_dashi(name):
                price = calc.dashi_pricing(qty)["avgUnitPrice"]  # 実質単価
            出.append({"name": name, "qty": qty, "price": price})
    return 出


def _合計(品):
    return int(round(sum(i["qty"] * i["price"] for i in 品)))


@owner_required
def order_create(request):
    if request.method == "POST":
        品 = _品を読む(request)
        if not 品:
            messages.error(request, "商品を1つ以上入れてください。")
        else:
            答 = orders.注文する({
                "manual": True, "payment": "手動入力",
                "date": request.POST.get("date", ""), "customerName": request.POST.get("customerName", ""),
                "memberId": request.POST.get("memberId", ""), "items": 品, "total": _合計(品),
                "status": _状態(request.POST.get("status")), "checked": request.POST.get("checked") == "1",
                "internalNote": request.POST.get("internalNote", ""),
            })
            if 答.get("status") == "ok":
                messages.success(request, "注文を作成しました")
                return redirect("manage:order_list")
            messages.error(request, 答.get("message") or "作れませんでした。")
    values = request.POST or {"date": timezone.localdate().isoformat(), "status": "受付中"}
    return render(request, "manage/order_form.html", {
        "order": None, "values": values, "items": _品を読む(request) if request.method == "POST" else [],
        "products": _商品候補(), "statuses": STATUSES, "dashi_names": list(calc.DASHI_NAMES),
    })


@owner_required
def order_edit(request, order_id: str):
    元 = [o for o in _まとめ() if o["orderId"] == order_id]
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
                "items": 品, "total": _合計(品),
                "status": _状態(request.POST.get("status")), "checked": request.POST.get("checked") == "1",
                "internalNote": request.POST.get("internalNote", ""),
            })
            if 答.get("status") == "ok":
                messages.success(request, "受取確認済み — 一覧から削除しました" if request.POST.get("checked") == "1" else "注文情報を更新しました")
                return redirect("manage:order_list")
            messages.error(request, 答.get("message") or "更新できませんでした。")
    values = request.POST or {
        "date": timezone.localtime(o["date"]).strftime("%Y-%m-%d") if o["date"] else "",
        "customerName": o["customerName"], "memberId": o["memberId"], "status": o["status"],
        "checked": "1" if o["checked"] else "", "internalNote": o["note"],
    }
    return render(request, "manage/order_form.html", {
        "order": o, "values": values, "items": _品を読む(request) if request.method == "POST" else o["items"],
        "products": _商品候補(), "statuses": STATUSES, "dashi_names": list(calc.DASHI_NAMES),
    })


@owner_required
@require_POST
def order_delete(request, order_id: str):
    答 = orders.注文を消す({"orderIds": [order_id]})
    if 答.get("status") == "ok":
        messages.success(request, "注文を削除しました")
    else:
        messages.error(request, 答.get("message") or "削除できませんでした。")
    return redirect("manage:order_list")
