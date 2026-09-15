"""データ分析と、その元になる売上の記録（メニュー収益・商品収益）。

集計は gasapi/analytics.集計()（GAS の getAnalyticsData と同じ形）。
記録の読み書きは gasapi/revenue.py（GAS の転送先と同じ）。
"""

import json

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Product
from apps.gasapi import analytics, revenue
from apps.records.models import RevenueRecord

from .permissions import owner_required

KINDS = {
    "menu": {"kind": RevenueRecord.MENU, "label": "メニュー収益", "name_key": "menuType", "qty_key": "count",
             "save": revenue.メニューを保存, "delete": revenue.メニューを消す, "list": revenue.メニューの記録},
    "product": {"kind": RevenueRecord.PRODUCT, "label": "商品収益", "name_key": "productName", "qty_key": "qty",
                "save": revenue.商品を保存, "delete": revenue.商品を消す, "list": revenue.商品の記録},
}


def _月の行(鍵, 月):
    return {
        "key": 鍵,
        "label": f"{鍵[:4]}年{int(鍵[5:7])}月" if len(鍵) >= 7 else 鍵,
        "product_sales": 月["sales"], "menu_sales": 月["menuRevenueTotal"],
        "revenue": 月["combinedRevenue"], "cost": 月["combinedCost"], "profit": 月["combinedProfit"],
        "menu_count": sum(月["menuCount"].values()),
    }


@owner_required
def analytics_view(request):
    集計 = analytics.集計()
    months = 集計["months"]
    rows = [_月の行(k, 集計["matrix"][k]) for k in months]
    selected = request.GET.get("month") or (months[0] if months else "")
    月 = 集計["matrix"].get(selected) if selected else None
    detail = None
    if 月:
        menu_rows = []
        for 種別 in 集計["menuTypes"]:
            内訳 = sorted(月["menuBreakdown"][種別].values(), key=lambda b: -b["unitPrice"])
            menu_rows.append({
                "type": 種別, "count": 月["menuCount"][種別], "revenue": 月["menuRevenue"][種別],
                "cost": 月["menuCost"][種別], "profit": 月["menuProfit"][種別], "breakdown": 内訳,
            })
        product_rows = sorted(月["productDetails"].values(), key=lambda p: (-p["sales"], p["name"]))
        detail = {
            "key": selected, "label": _月の行(selected, 月)["label"], "menu_rows": menu_rows,
            "product_rows": product_rows, "totals": _月の行(selected, 月),
        }
    return render(request, "manage/analytics.html", {
        "rows": rows, "detail": detail, "selected": selected,
        "totals": {
            "revenue": sum(r["revenue"] for r in rows), "cost": sum(r["cost"] for r in rows),
            "profit": sum(r["profit"] for r in rows),
        },
    })


@owner_required
def revenue_list(request, kind: str):
    設定 = KINDS.get(kind)
    if not 設定:
        return redirect("manage:analytics")
    records = 設定["list"]()["records"]
    months = sorted({r["date"][:7] for r in records if r.get("date")}, reverse=True)
    month = request.GET.get("month", "")
    if month:
        records = [r for r in records if (r.get("date") or "").startswith(month)]
    records = sorted(records, key=lambda r: (r.get("date") or "", r["rowIdx"]), reverse=True)
    edit_row = request.GET.get("edit")
    editing = next((r for r in records if str(r["rowIdx"]) == str(edit_row)), None) if edit_row else None
    return render(request, "manage/revenue_list.html", {
        "kind": kind, "label": 設定["label"], "records": records, "months": months, "month": month,
        "menu_types": revenue.メニュー種別, "editing": editing,
        "products": list(Product.objects.filter(deleted=False).order_by("sheet_row").values_list("name", "price")),
        "today": timezone.localdate().isoformat(),
        "total": {
            "amount": sum(r["totalAmount"] for r in records), "cost": sum(r["totalCost"] for r in records),
            "profit": sum(r["profit"] for r in records),
        },
    })


@owner_required
@require_POST
def revenue_save(request, kind: str):
    設定 = KINDS.get(kind)
    if not 設定:
        return redirect("manage:analytics")
    date = request.POST.get("date", "")
    row = request.POST.get("rowIdx", "")
    if row:
        d = {"rowIdx": row, "date": date, 設定["name_key"]: request.POST.get("name", ""),
             設定["qty_key"]: request.POST.get("qty", "1"), "unitPrice": request.POST.get("unitPrice", "0"),
             "unitCost": request.POST.get("unitCost", "0"), "note": request.POST.get("note", "")}
    else:
        try:
            並び = json.loads(request.POST.get("rows_json") or "[]")
        except ValueError:
            並び = []
        並び = [x for x in 並び if isinstance(x, dict) and str(x.get("name") or "").strip()]
        if not 並び:
            messages.error(request, "記録を1行以上入れてください。")
            return redirect(f"/manage/revenue/{kind}/")
        d = {"date": date, "records": [{
            "date": date, 設定["name_key"]: x["name"], 設定["qty_key"]: x.get("qty", 1),
            "unitPrice": x.get("unitPrice", 0), "unitCost": x.get("unitCost", 0), "note": x.get("note", ""),
        } for x in 並び]}
    答 = 設定["save"](d)
    if 答.get("status") == "ok":
        messages.success(request, f"{設定['label']}を保存しました（追加 {答.get('createdCount', 0)}・更新 {答.get('updatedCount', 0)}）。")
    else:
        messages.error(request, 答.get("message") or "保存できませんでした。")
    return redirect(f"/manage/revenue/{kind}/?month={date[:7]}" if date else f"/manage/revenue/{kind}/")


@owner_required
@require_POST
def revenue_delete(request, kind: str, row: int):
    設定 = KINDS.get(kind)
    if not 設定:
        return redirect("manage:analytics")
    答 = 設定["delete"]({"rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, "記録を削除しました。")
    else:
        messages.error(request, 答.get("message") or "削除できませんでした。")
    return redirect(request.POST.get("next") or f"/manage/revenue/{kind}/")
