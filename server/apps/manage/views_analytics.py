"""データ分析と、その元になる売上の記録（メニュー収益・商品収益）。

**中身は旧管理アプリ（admin/index.html の #page-analytics）と同じ。**
集計は analytics_calc.build()（旧アプリの計算の写し）。記録の読み書きは gasapi/revenue.py。
"""

import json
from collections import OrderedDict

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.utils.safestring import mark_safe
from django.views.decorators.http import require_POST

from apps.content.models import Product
from apps.gasapi import revenue
from apps.gasapi.admin_product import _原価の表
from apps.records.models import RevenueRecord

from . import analytics_calc as calc
from .permissions import owner_required

KINDS = {
    "menu": {"kind": RevenueRecord.MENU, "label": "メニュー収益", "name_key": "menuType", "qty_key": "count",
             "save": revenue.メニューを保存, "delete": revenue.メニューを消す, "list": revenue.メニューの記録},
    "product": {"kind": RevenueRecord.PRODUCT, "label": "商品収益", "name_key": "productName", "qty_key": "qty",
                "save": revenue.商品を保存, "delete": revenue.商品を消す, "list": revenue.商品の記録},
}
# メニュー原価・メニュー粗利の列は出さない（院長の希望 2026-09-16）
FIN_COLUMNS = [("combinedRevenue", "総売上(円)"), ("combinedCost", "総原価(円)"), ("combinedProfit", "総粗利(円)"),
               ("sales", "物販売上(円)"), ("cost", "物販原価(円)"), ("profit", "物販粗利(円)"),
               ("menuRevenueTotal", "メニュー売上(円)")]


def _label(key: str) -> str:
    return f"{key[:4]}/{key[5:7]}" if len(key) >= 7 else key


@owner_required
def analytics_view(request):
    res = calc.build()
    matrix = res["matrix"]
    months = [m for m in res["months"] if calc.month_visible(m)]  # 新しい順
    ordered = list(reversed(months))  # グラフは古い順
    latest = months[0] if months else ""

    # a) 月別比較
    comparison_rows = []
    for m in months:
        cur = matrix.get(m, {})
        prev = matrix.get(calc.shift_month(m, -1), {})
        last_year = matrix.get(calc.shift_month(m, -12), {})
        comparison_rows.append({
            "month": m,
            "revenue": calc.yen(cur.get("combinedRevenue")),
            "revenue_mom": calc.comparison(cur.get("combinedRevenue"), prev.get("combinedRevenue")),
            "revenue_yoy": calc.comparison(cur.get("combinedRevenue"), last_year.get("combinedRevenue")),
            "profit": calc.yen(cur.get("combinedProfit")),
            "profit_mom": calc.comparison(cur.get("combinedProfit"), prev.get("combinedProfit")),
            "profit_yoy": calc.comparison(cur.get("combinedProfit"), last_year.get("combinedProfit")),
        })

    # b) 粗利サマリー
    stats = matrix.get(latest) if latest else None
    stats = stats or calc._empty_month()
    gross_cards = [
        (f"{latest} 総売上", calc.yen(stats["combinedRevenue"])),
        (f"{latest} 総粗利", calc.yen(stats["combinedProfit"])),
        ("物販粗利率", calc.ratio_label(stats["profit"], stats["sales"])),
        # メニュー粗利率は出さない（院長の希望 2026-09-16）
        ("物販粗利", calc.yen(stats["profit"])),
        ("メニュー粗利", calc.yen(stats["menuProfitTotal"])),
    ] if latest else []

    # c) 月別データグラフ（4本）
    menu_names = sorted({n for m in ordered for n in matrix[m]["menuDetails"]})
    product_names = sorted({n for m in ordered for n in matrix[m]["products"]})
    menu_target = request.GET.get("menu_target", "")
    product_target = request.GET.get("product_target", "")
    if menu_target not in menu_names:
        menu_target = ""
    if product_target not in product_names:
        product_target = ""

    def series(getter):
        return [{"label": _label(m), "value": getter(matrix[m])} for m in ordered]

    charts = [
        {"title": "メニュー売上", "control": "menu_target", "options": menu_names, "selected": menu_target,
         "all_label": "全メニュー合計",
         "subtitle": (f"「{menu_target}」の月ごとの売上" if menu_target else "月ごとのメニュー売上合計"),
         "svg": calc.bar_chart_svg(series(lambda s: (s["menuDetails"].get(menu_target, {}).get("revenue", 0) if menu_target else s["menuRevenueTotal"])), "#5a7a4e")},
        {"title": "物販売上", "control": "product_target", "options": product_names, "selected": product_target,
         "all_label": "全商品合計",
         "subtitle": (f"「{product_target}」の月ごとの売上" if product_target else "月ごとの物販売上合計"),
         "svg": calc.bar_chart_svg(series(lambda s: (s["products"].get(product_target, {}).get("sales", 0) if product_target else s["sales"])), "#b8752e")},
        {"title": "総売上", "subtitle": "メニュー売上と物販売上を合計した月別推移",
         "svg": calc.bar_chart_svg(series(lambda s: s["combinedRevenue"]), "#1a4a73")},
        {"title": "総粗利", "subtitle": "原価を引いた後の月別粗利推移",
         "svg": calc.bar_chart_svg(series(lambda s: s["combinedProfit"]), "#2e7d32")},
    ]
    for c in charts:
        c["svg"] = mark_safe(c["svg"])

    # d) 収益サマリー（月別）
    profit_rows = [{"month": m, "cells": [calc.num(matrix[m].get(k, 0)) for k, _ in FIN_COLUMNS]} for m in months]

    # e) メニュー別収益（月別）
    menu_rows = []
    for m in months:
        s = matrix[m]
        cells = []
        for t in res["menuTypes"]:
            breakdown = sorted(s["menuBreakdown"][t].values(), key=lambda b: b["unitPrice"])
            cells.append({
                "amount": calc.yen(s["menuRevenue"][t]) if s["menuRevenue"][t] else "-",
                "count": s["menuCount"][t], "cost": calc.yen(s["menuCost"][t]), "profit": calc.yen(s["menuProfit"][t]),
                "breakdown": [{"unit": f"¥{b['unitPrice']:,}", "revenue": calc.yen(b["revenue"]),
                               "profit": calc.yen(b["profit"]), "count": b["count"]} for b in breakdown],
            })
        menu_rows.append({"month": m, "cells": cells,
                          "total": {"amount": calc.yen(s["menuRevenueTotal"]), "cost": calc.yen(s["menuCostTotal"]),
                                    "profit": calc.yen(s["menuProfitTotal"]), "count": sum(s["menuCount"].values())}})

    # j/k) 月を選んだ詳細
    detail_month = request.GET.get("month", "")
    if detail_month not in months:
        detail_month = ""
    product_detail_rows, menu_detail_rows = [], []
    if detail_month:
        s = matrix[detail_month]
        details = sorted(s["productDetails"].values(), key=lambda d: (d["name"], d["sortOrder"], -d["price"]))
        for d in details:
            pm = res["master"].get(d["name"])
            product_detail_rows.append({
                "name": d["name"], "list_price": (f"¥{pm.price:,}" if pm and pm.price else "-"),
                "unit_cost": (f"¥{int(d['cost'] // d['qty']):,}" if d["qty"] else "-"),
                "price_label": d["priceLabel"], "qty": f"{d['qty']}個",
                "sales": f"¥{int(d['sales']):,}", "profit": f"¥{int(d['profit']):,}",
            })
        for name in sorted(s["menuDetails"]):
            d = s["menuDetails"][name]
            menu_detail_rows.append({
                "name": name, "count": d["count"], "revenue": calc.yen(d["revenue"]), "cost": calc.yen(d["cost"]),
                "profit": calc.yen(d["profit"]),
                "breakdown": [{"unit": f"¥{b['unitPrice']:,}", "count": b["count"], "revenue": calc.yen(b["revenue"]),
                               "profit": calc.yen(b["profit"])} for b in sorted(d["breakdown"].values(), key=lambda b: b["unitPrice"])],
            })

    # l) 商品別販売数（推移）
    display_products = sorted({*res["master"].keys(), *res["products"],
                               *[n for m in months for n in matrix[m]["products"]]})
    qty_rows = [{"month": m, "cells": [(matrix[m]["products"].get(p, {}).get("qty") or "-") for p in display_products]}
                for m in months]

    return render(request, "manage/analytics.html", {
        "months": months, "latest": latest, "no_data": not months,
        "comparison_rows": comparison_rows, "gross_cards": gross_cards, "charts": charts,
        "fin_columns": [c for _, c in FIN_COLUMNS], "profit_rows": profit_rows,
        "menu_types": res["menuTypes"], "menu_rows": menu_rows,
        "detail_month": detail_month, "product_detail_rows": product_detail_rows, "menu_detail_rows": menu_detail_rows,
        "display_products": display_products, "qty_rows": qty_rows,
        "query": {k: v for k, v in request.GET.items() if k in ("menu_target", "product_target", "month")},
    })


# ── 売上の記録（g/h/i/f）──

def _grouped(records: list) -> list:
    """登録済み一覧: 日付で束ね、新しい日から。最初の束だけ開く。"""
    by_date = OrderedDict()
    for r in sorted(records, key=lambda x: (x.get("date") or "", x["rowIdx"]), reverse=True):
        by_date.setdefault((r.get("date") or "")[:10], []).append(r)
    out = []
    for i, (d, rows) in enumerate(by_date.items()):
        qty_key = "count" if "menuType" in rows[0] else "qty"
        name_key = "menuType" if "menuType" in rows[0] else "productName"
        chips = OrderedDict()
        for r in rows:
            c = chips.setdefault(r.get(name_key, ""), {"count": 0, "total": 0})
            c["count"] += r.get(qty_key, 0)
            c["total"] += r.get("totalAmount", 0)
        out.append({
            "date": d.replace("-", "/"), "rows": rows, "open": i == 0, "n": len(rows),
            "day_count": sum(r.get(qty_key, 0) for r in rows),
            "day_total": sum(r.get("totalAmount", 0) for r in rows),
            "day_profit": sum(r.get("profit", 0) for r in rows),
            "chips": [{"name": k, **v} for k, v in chips.items()],
        })
    return out


@owner_required
def revenue_list(request, kind: str):
    設定 = KINDS.get(kind)
    if not 設定:
        return redirect("manage:analytics")
    records = 設定["list"]()["records"]
    for r in records:
        if kind == "product" and calc.is_dashi(r.get("productName", "")):
            r["dashi"] = calc.dashi_pricing(r.get("qty", 0))
    months = sorted({r["date"][:7] for r in records if r.get("date")}, reverse=True)
    month = request.GET.get("month", "")
    shown = [r for r in records if (r.get("date") or "").startswith(month)] if month else records
    edit_row = request.GET.get("edit")
    editing = next((r for r in records if str(r["rowIdx"]) == str(edit_row)), None) if edit_row else None
    products = list(Product.objects.filter(deleted=False).order_by("sheet_row").values("name", "price"))
    cost_map = _原価の表()
    return render(request, "manage/revenue_list.html", {
        "kind": kind, "label": 設定["label"], "groups": _grouped(shown), "months": months, "month": month,
        "menu_types": revenue.メニュー種別, "editing": editing,
        # json_script で埋める（商品名に "</script>" が入っても壊れない）。
        "products_data": [{"name": p["name"], "price": p["price"] or 0, "cost": cost_map.get(p["name"], 0)} for p in products],
        "product_names": {p["name"] for p in products},
        "today": timezone.localdate().isoformat(),
        "meta": {"count": len(shown), "amount": sum(r["totalAmount"] for r in shown), "profit": sum(r["profit"] for r in shown)},
        "dashi_names": list(calc.DASHI_NAMES), "dashi_cost": calc.DASHI_DEFAULT_COST,
    })


@owner_required
@require_POST
def revenue_save(request, kind: str):
    設定 = KINDS.get(kind)
    if not 設定:
        return redirect("manage:analytics")
    date = request.POST.get("date", "")
    row = request.POST.get("rowIdx", "")

    def _dashi_price(name, qty, unit_price):
        # 「実質単価 (自動)」: だしは個数で単価が決まる
        if kind == "product" and calc.is_dashi(name):
            return calc.dashi_pricing(int(float(qty or 0)))["avgUnitPrice"]
        return unit_price

    if row:
        name = request.POST.get("name", "")
        qty = request.POST.get("qty", "1")
        d = {"rowIdx": row, "date": date, 設定["name_key"]: name, 設定["qty_key"]: qty,
             "unitPrice": _dashi_price(name, qty, request.POST.get("unitPrice", "0")),
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
            "unitPrice": _dashi_price(x["name"], x.get("qty", 1), x.get("unitPrice", 0)),
            "unitCost": x.get("unitCost", 0), "note": x.get("note", ""),
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
