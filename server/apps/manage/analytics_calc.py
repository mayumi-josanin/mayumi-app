"""データ分析の集計。**旧管理アプリ（admin/index.html）の計算をそのまま写す。**

旧アプリの決めごと:
- **ショップの注文は集計に入れない。**「商品収益を記録」で入れた分だけが
  商品別販売数・商品別収益詳細・収益サマリーに入る（画面の注記にそう書いてある）。
- メニュー収益は種別（母乳外来／ビジリス／教室／その他）と単価ごとにまとめる。
  「その他」で入れた固有の名前は menuDetails に個別に残す。
- 天然だし調味粉は個数で割引の段階が決まる（2個まで20%、3個は2+1、4個は3+1、5個は4+1、6個以上は全部25%）。
  価格帯ごとに分けて「商品別収益詳細」に出す。
- 表に出す月は 2026年4月以降だけ。

数字の丸め・並び・ラベルは旧アプリの関数名をコメントに残してある。
"""

from datetime import date
from decimal import Decimal

from apps.content.models import Product
from apps.records.models import RevenueRecord

MENU_TYPES = ["母乳外来", "ビジリス", "教室", "その他"]  # MENU_REVENUE_TYPES
DASHI_NAMES = ("天然だし調味粉", "天然だし調理粉")
DASHI_TIERS = [("20", 2380, "2,380円(20%OFF)"), ("25", 2235, "2,235円(25%OFF)"),
               ("30", 2086, "2,086円(30%OFF)"), ("35", 1937, "1,937円(35%OFF)")]
DASHI_TIER_PRICE = {k: p for k, p, _ in DASHI_TIERS}
DASHI_DEFAULT_COST = 1490
DASHI_LIST_PRICE = 2980


def _f(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, Decimal):
        return float(v)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def is_dashi(name: str) -> bool:
    return (name or "").strip() in DASHI_NAMES


def dashi_tiers(count: int) -> dict:
    """buildDashiTierBreakdown。"""
    count = max(0, int(count or 0))
    b = {"20": 0, "25": 0, "30": 0, "35": 0}
    if count <= 0:
        return b
    if count <= 2:
        b["20"] = count
    elif count == 3:
        b["20"], b["25"] = 2, 1
    elif count == 4:
        b["20"], b["30"] = 3, 1
    elif count == 5:
        b["20"], b["35"] = 4, 1
    else:
        b["25"] = count
    return b


def dashi_pricing(qty: int) -> dict:
    """calculateDashiPricing: totalRevenue / avgUnitPrice / breakdown / breakdownText。"""
    count = max(0, int(qty or 0))
    if count <= 0:
        return {"totalRevenue": 0, "avgUnitPrice": DASHI_LIST_PRICE, "breakdown": dashi_tiers(0), "text": ""}
    b = dashi_tiers(count)
    total = sum(DASHI_TIER_PRICE[k] * n for k, n in b.items())
    text = " / ".join(f"{label} × {b[k]}個" for k, _, label in DASHI_TIERS if b[k])
    return {"totalRevenue": total, "avgUnitPrice": total / count, "breakdown": b, "text": text}


def month_key(d: date | None) -> str:
    return f"{d.year}-{d.month:02d}" if d else ""


def shift_month(key: str, diff: int) -> str:
    """shiftMonthKey。"""
    try:
        y, m = int(key[:4]), int(key[5:7])
    except (ValueError, TypeError):
        return ""
    idx = y * 12 + (m - 1) + diff
    return f"{idx // 12}-{idx % 12 + 1:02d}"


def month_visible(key: str) -> bool:
    """2026年4月以降だけ出す（renderAnalytics の filteredMonths）。形が違う鍵はそのまま出す。"""
    try:
        y, m = int(key[:4]), int(key[5:7])
    except (ValueError, TypeError):
        return True
    return y > 2026 or (y == 2026 and m >= 4)


def unit_price_label(unit_price: float, dashi: bool) -> str:
    """formatAnalyticsUnitPriceLabel。"""
    if not unit_price or unit_price <= 0:
        return "-"
    if float(unit_price).is_integer():
        text = f"{int(unit_price):,}"
    else:
        text = f"{unit_price:,.2f}".rstrip("0").rstrip(".")
    return f"実質 ¥{text}" if dashi else f"¥{text}"


def yen(v) -> str:
    return f"¥{int(round(_f(v))):,}"


def num(v) -> str:
    return f"{int(round(_f(v))):,}"


def comparison(current, base) -> dict:
    """formatComparisonRate。rate は % 。base が 0 なら None。"""
    c, b = _f(current), _f(base)
    if b == 0:
        return {"rate": None, "text": "-", "cls": "muted"}
    rate = (c - b) / b * 100
    sign = "+" if rate > 0 else ""
    return {"rate": rate, "text": f"{sign}{rate:.1f}%", "cls": "up" if rate > 0 else ("down" if rate < 0 else "flat")}


def ratio_label(numerator, denominator) -> str:
    """formatRatioLabel。"""
    d = _f(denominator)
    if d <= 0:
        return "-"
    return f"{_f(numerator) / d * 100:.1f}%"


def _empty_month() -> dict:
    return {
        "products": {}, "productDetails": {}, "sales": 0.0, "cost": 0.0, "profit": 0.0,
        "menuRevenue": {t: 0.0 for t in MENU_TYPES}, "menuCost": {t: 0.0 for t in MENU_TYPES},
        "menuProfit": {t: 0.0 for t in MENU_TYPES}, "menuCount": {t: 0 for t in MENU_TYPES},
        "menuBreakdown": {t: {} for t in MENU_TYPES}, "menuDetails": {},
        "menuRevenueTotal": 0.0, "menuCostTotal": 0.0, "menuProfitTotal": 0.0,
        "combinedRevenue": 0.0, "combinedCost": 0.0, "combinedProfit": 0.0,
    }


def record_totals(r) -> dict:
    """getProductRevenueRecordTotals / メニューも同じ式。"""
    qty = max(0, int(r.quantity or 0))
    unit_price = max(0.0, _f(r.unit_price))
    unit_cost = max(0.0, _f(r.unit_cost))
    total = qty * unit_price
    cost = qty * unit_cost
    return {"qty": qty, "unitPrice": unit_price, "unitCost": unit_cost,
            "totalAmount": total, "totalCost": cost, "totalProfit": total - cost}


def build() -> dict:
    """buildAnalyticsWithManualProductRevenue ＋ GAS の getAnalyticsData のメニュー部分。"""
    master = {p.name.strip(): p for p in Product.objects.filter(deleted=False)}
    matrix: dict[str, dict] = {}

    def month(key):
        return matrix.setdefault(key, _empty_month())

    product_names = set()
    for r in RevenueRecord.objects.filter(kind=RevenueRecord.PRODUCT, deleted=False):
        key = month_key(r.recorded_on)
        name = (r.name or "").strip()
        if not key or not name:
            continue
        t = record_totals(r)
        if t["qty"] <= 0:
            continue
        m = month(key)
        p = m["products"].setdefault(name, {"qty": 0, "sales": 0.0, "cost": 0.0, "profit": 0.0})
        p["qty"] += t["qty"]
        p["sales"] += t["totalAmount"]
        p["cost"] += t["totalCost"]
        p["profit"] += t["totalProfit"]
        m["sales"] += t["totalAmount"]
        m["cost"] += t["totalCost"]
        m["profit"] += t["totalProfit"]
        product_names.add(name)

        pm = master.get(name)
        sort_order = pm.sort_order if (pm and pm.sort_order is not None) else 1000
        dashi = is_dashi(name)
        entries = []
        if dashi:
            # 価格帯ごとに分ける（parseProductRevenuePriceBreakdownEntries）
            for k, price, label in DASHI_TIERS:
                n = dashi_tiers(t["qty"])[k]
                if n:
                    entries.append({"price": price, "priceLabel": f"¥{price:,} ({label.split('(')[1]}",
                                    "qty": n, "sales": price * n, "cost": t["unitCost"] * n,
                                    "profit": price * n - t["unitCost"] * n})
        if not entries:
            entries = [{"price": t["unitPrice"], "priceLabel": unit_price_label(t["unitPrice"], dashi),
                        "qty": t["qty"], "sales": t["totalAmount"], "cost": t["totalCost"], "profit": t["totalProfit"]}]
        for e in entries:
            dk = f"{name}::{e['priceLabel']}"
            cur = m["productDetails"].setdefault(dk, {
                "name": name, "price": e["price"], "priceLabel": e["priceLabel"],
                "qty": 0, "sales": 0.0, "cost": 0.0, "profit": 0.0, "sortOrder": sort_order})
            cur["qty"] += e["qty"]
            cur["sales"] += e["sales"]
            cur["cost"] += e["cost"]
            cur["profit"] += e["profit"]

    for r in RevenueRecord.objects.filter(kind=RevenueRecord.MENU, deleted=False):
        key = month_key(r.recorded_on)
        if not key:
            continue
        m = month(key)
        name = (r.name or "").strip()
        bucket = name if name in MENU_TYPES else "その他"
        t = record_totals(r)
        cnt, price, cost_u = t["qty"], t["unitPrice"], t["unitCost"]
        rev, cost, prof = t["totalAmount"], t["totalCost"], t["totalProfit"]
        m["menuRevenue"][bucket] += rev
        m["menuCost"][bucket] += cost
        m["menuProfit"][bucket] += prof
        m["menuCount"][bucket] += cnt
        pk = str(int(price))
        box = m["menuBreakdown"][bucket].setdefault(pk, {"count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "unitPrice": int(price)})
        box["count"] += cnt
        box["revenue"] += rev
        box["cost"] += cost
        box["profit"] += prof
        det = m["menuDetails"].setdefault(name or bucket, {"count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "breakdown": {}})
        det["count"] += cnt
        det["revenue"] += rev
        det["cost"] += cost
        det["profit"] += prof
        db = det["breakdown"].setdefault(pk, {"count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "unitPrice": int(price)})
        db["count"] += cnt
        db["revenue"] += rev
        db["cost"] += cost
        db["profit"] += prof
        m["menuRevenueTotal"] += rev
        m["menuCostTotal"] += cost
        m["menuProfitTotal"] += prof

    for m in matrix.values():
        m["combinedRevenue"] = m["menuRevenueTotal"] + m["sales"]
        m["combinedCost"] = m["menuCostTotal"] + m["cost"]
        m["combinedProfit"] = m["menuProfitTotal"] + m["profit"]

    months = sorted(matrix.keys(), reverse=True)
    return {
        "months": months, "matrix": matrix, "menuTypes": list(MENU_TYPES),
        "products": sorted(product_names), "master": master,
    }


# ── 棒グラフ（buildMonthlyAxisChartSvg をそのまま SVG 文字列で） ──

def chart_ceiling(value: float) -> float:
    """getMonthlyChartCeiling。"""
    import math

    v = _f(value)
    if v <= 0:
        return 1
    magnitude = 10 ** math.floor(math.log10(v))
    normalized = v / magnitude
    nice = 1 if normalized <= 1 else 2 if normalized <= 2 else 5 if normalized <= 5 else 10
    return nice * magnitude


def axis_label(value: float) -> str:
    """formatMonthlyChartAxisValue。"""
    v = _f(value)
    if abs(v) >= 100_000_000:
        u = v / 100_000_000
        return f"{int(u)}億" if u.is_integer() else f"{u:.1f}億"
    if abs(v) >= 10_000:
        u = v / 10_000
        return f"{int(u)}万" if u.is_integer() else f"{u:.1f}万"
    return f"{int(round(v)):,}"


def bar_chart_svg(points: list, color: str) -> str:
    """points = [{"label": "2026/04", "value": 12345}, ...]（古い月→新しい月）。"""
    from django.utils.html import escape

    if not points:
        return '<div class="monthly-chart-empty">表示できるデータがありません</div>'
    width = max(420, len(points) * 86 + 96)
    height = 280
    top, right, bottom, left = 18, 18, 52, 62
    inner_w = width - left - right
    inner_h = height - top - bottom
    max_value = max(_f(p["value"]) for p in points)
    ceiling = chart_ceiling(max_value)
    step = inner_w / len(points)
    bar_w = min(38, max(18, step * 0.52))
    ticks = 4
    parts = [f'<svg class="monthly-chart-svg" viewBox="0 0 {width} {height}" width="{width}" height="{height}" role="img" aria-label="月別棒グラフ">',
             f'<line x1="{left}" y1="{top}" x2="{left}" y2="{top + inner_h}" stroke="#cdbca8" stroke-width="1.5" />',
             f'<line x1="{left}" y1="{top + inner_h}" x2="{width - right}" y2="{top + inner_h}" stroke="#cdbca8" stroke-width="1.5" />']
    for i in range(ticks + 1):
        ratio = i / ticks
        y = top + inner_h - inner_h * ratio
        parts.append(f'<g><line x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" stroke="#e6ddd2" stroke-width="1" />'
                     f'<text x="{left - 8}" y="{y + 4:.1f}" text-anchor="end" font-size="10" fill="#9a8070">{escape(axis_label(ceiling * ratio))}</text></g>')
    for i, p in enumerate(points):
        v = _f(p["value"])
        bar_h = (v / ceiling) * inner_h if ceiling > 0 else 0
        x = left + step * i + (step - bar_w) / 2
        y = top + inner_h - bar_h
        label_y = max(y - 6, 12)
        parts.append(f'<g><rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{bar_h:.1f}" rx="6" ry="6" fill="{color}" />'
                     f'<text x="{x + bar_w / 2:.1f}" y="{label_y:.1f}" text-anchor="middle" font-size="10" font-weight="700" fill="#3d2e1e">¥{int(round(v)):,}</text>'
                     f'<text x="{x + bar_w / 2:.1f}" y="{height - 18}" text-anchor="middle" font-size="11" fill="#6b5040">{escape(str(p["label"]))}</text></g>')
    parts.append(f'<text x="{width / 2}" y="{height - 2}" text-anchor="middle" font-size="11" fill="#6b5040">月</text>')
    parts.append(f'<text x="16" y="{height / 2}" text-anchor="middle" font-size="11" fill="#6b5040" transform="rotate(-90 16 {height / 2})">金額 (円)</text>')
    parts.append("</svg>")
    return '<div class="monthly-axis-chart">' + "".join(parts) + "</div>"
