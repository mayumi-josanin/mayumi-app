"""分析画面。**GAS の getAnalyticsData（6593行・292行）と同じ計算。**

計画は [注文の表を移す](docs/design/注文の表を移す.md)。

## 3つの表を読む

    注文（OrderLine）           商品の販売。**いまは0件**
    売上・メニュー（RevenueRecord kind=menu）    112件
    売上・商品（RevenueRecord kind=product）      65件

**注文だけを移して分析を残すと、分析が注文を0件として計算する。**
だから注文と一緒に移す。（お知らせを移してカテゴリが食い違ったのと同じ形）

## 天然だしの値引き

個数によって単価が変わる。**この規則は GAS と1文字も違えられない。**
違えると粗利の数字が静かにずれる。

    1〜2個   20%OFF ×個数
    3個      20%OFF ×2 ＋ 25%OFF ×1
    4個      20%OFF ×3 ＋ 30%OFF ×1
    5個      20%OFF ×4 ＋ 35%OFF ×1
    6個以上  25%OFF ×個数
"""

from decimal import Decimal

from apps.records.models import OrderLine, RevenueRecord

メニュー種別 = ["母乳外来", "ビジリス", "教室", "その他"]

だしの名前 = ["天然だし調味粉", "天然だし調理粉"]

だしの段 = [
    {"key": "20", "label": "2,380円(20%OFF)", "price": 2380, "sortOrder": 1},
    {"key": "25", "label": "2,235円(25%OFF)", "price": 2235, "sortOrder": 2},
    {"key": "30", "label": "2,086円(30%OFF)", "price": 2086, "sortOrder": 3},
    {"key": "35", "label": "1,937円(35%OFF)", "price": 1937, "sortOrder": 4},
]


def _数(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _整数(v):
    try:
        return int(float(v or 0))
    except (TypeError, ValueError):
        return 0


def _だしか(名):
    return str(名 or "").strip() in だしの名前


def だしの内訳(個数):
    """GAS の buildDashiTierBreakdown_ と**同じ規則。**"""
    n = max(0, _整数(個数))
    出 = {"20": 0, "25": 0, "30": 0, "35": 0}
    if n <= 0:
        return 出
    if n <= 2:
        出["20"] = n
        return 出
    if n == 3:
        出["20"], 出["25"] = 2, 1
        return 出
    if n == 4:
        出["20"], 出["30"] = 3, 1
        return 出
    if n == 5:
        出["20"], 出["35"] = 4, 1
        return 出
    出["25"] = n
    return 出


def _だしの鍵(商品名, 段の鍵):
    """GAS の getDashiDetailKey_ と同じ。"""
    return f"{str(商品名 or '').strip()}|dashi|{段の鍵}"


def _月(日付):
    return f"{日付.year}-{日付.month:02d}" if 日付 else ""


def _空の月():
    return {
        "products": {},
        "sales": 0.0,
        "cost": 0.0,
        "profit": 0.0,
        "menuRevenue": {t: 0.0 for t in メニュー種別},
        "menuCost": {t: 0.0 for t in メニュー種別},
        "menuProfit": {t: 0.0 for t in メニュー種別},
        "menuCount": {t: 0 for t in メニュー種別},
        "menuBreakdown": {t: {} for t in メニュー種別},
        "menuDetails": {},
        "menuRevenueTotal": 0.0,
        "menuCostTotal": 0.0,
        "menuProfitTotal": 0.0,
        "combinedRevenue": 0.0,
        "combinedCost": 0.0,
        "combinedProfit": 0.0,
        "combinedSales": 0.0,
        "productDetails": {},
    }


def _商品を足す(月, 名, 個数, 売上, 原価計, 粗利):
    箱 = 月["products"].setdefault(名, {"qty": 0, "sales": 0.0, "cost": 0.0, "profit": 0.0})
    箱["qty"] += 個数
    箱["sales"] += 売上
    箱["cost"] += 原価計
    箱["profit"] += 粗利
    月["sales"] += 売上
    月["cost"] += 原価計
    月["profit"] += 粗利


def _内訳を足す(月, 鍵, 名, 単価, 表示, 個数, 売上, 原価計, 粗利, 並び=0):
    箱 = 月["productDetails"].setdefault(鍵, {
        "name": 名, "price": 単価, "priceLabel": 表示,
        "qty": 0, "sales": 0.0, "cost": 0.0, "profit": 0.0, "sortOrder": 並び,
    })
    箱["qty"] += 個数
    箱["sales"] += 売上
    箱["cost"] += 原価計
    箱["profit"] += 粗利


def 集計():
    """GAS の getAnalyticsData と同じ形で返す。"""
    月表 = {}

    def 月を出す(鍵):
        if 鍵 not in 月表:
            月表[鍵] = _空の月()
        return 月表[鍵]

    # ── ① 注文（商品の販売）──
    for r in OrderLine.objects.all():
        if (r.status or "").strip() == "キャンセル済":
            continue
        if not r.ordered_at:
            continue
        月 = 月を出す(_月(r.ordered_at.date()))
        名 = (r.product_name or "").strip()
        個数 = _整数(r.quantity)
        単価 = _数(r.unit_price)
        原価 = _数(r.cost_price)
        _商品を足す(月, 名, 個数, _数(r.subtotal), 原価 * 個数, _数(r.profit))
        _だしか含めて内訳(月, 名, 個数, 単価, 原価, _数(r.subtotal), _数(r.profit))

    # ── ② 売上・メニュー ──
    # **kind の値は「メニュー」「商品」。**`"menu"` で絞って0件になった
    # （2026-09-07）。**モデルの定数を使う。**文字を書き写さない。
    for r in RevenueRecord.objects.filter(kind=RevenueRecord.MENU, deleted=False):
        if not r.recorded_on:
            continue
        月 = 月を出す(_月(r.recorded_on))
        種別 = (r.name or "").strip()
        まとめ = 種別 if 種別 in メニュー種別 else "その他"
        件数 = _整数(r.quantity)
        単価 = _数(r.unit_price)
        原価 = _数(r.unit_cost)
        売上 = 単価 * 件数
        原価計 = 原価 * 件数
        粗利 = 売上 - 原価計

        月["menuRevenue"][まとめ] += 売上
        月["menuCost"][まとめ] += 原価計
        月["menuProfit"][まとめ] += 粗利
        月["menuCount"][まとめ] += 件数

        単価鍵 = str(_整数(単価))
        箱 = 月["menuBreakdown"][まとめ].setdefault(
            単価鍵, {"count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "unitPrice": _整数(単価)})
        箱["count"] += 件数
        箱["revenue"] += 売上
        箱["cost"] += 原価計
        箱["profit"] += 粗利

        詳 = 月["menuDetails"].setdefault(種別, {
            "count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "breakdown": {}})
        詳["count"] += 件数
        詳["revenue"] += 売上
        詳["cost"] += 原価計
        詳["profit"] += 粗利
        詳箱 = 詳["breakdown"].setdefault(
            単価鍵, {"count": 0, "revenue": 0.0, "cost": 0.0, "profit": 0.0, "unitPrice": _整数(単価)})
        詳箱["count"] += 件数
        詳箱["revenue"] += 売上
        詳箱["cost"] += 原価計
        詳箱["profit"] += 粗利

        月["menuRevenueTotal"] += 売上
        月["menuCostTotal"] += 原価計
        月["menuProfitTotal"] += 粗利

    # ── ③ 売上・商品 ──
    for r in RevenueRecord.objects.filter(kind=RevenueRecord.PRODUCT, deleted=False):
        if not r.recorded_on:
            continue
        月 = 月を出す(_月(r.recorded_on))
        名 = (r.name or "").strip()
        個数 = _整数(r.quantity)
        単価 = _数(r.unit_price)
        原価 = _数(r.unit_cost)
        売上 = 単価 * 個数
        原価計 = 原価 * 個数
        _商品を足す(月, 名, 個数, 売上, 原価計, 売上 - 原価計)
        _だしか含めて内訳(月, 名, 個数, 単価, 原価, 売上, 売上 - 原価計)

    # ── まとめ ──
    for 月 in 月表.values():
        月["combinedRevenue"] = 月["menuRevenueTotal"] + 月["sales"]
        月["combinedCost"] = 月["menuCostTotal"] + 月["cost"]
        月["combinedProfit"] = 月["menuProfitTotal"] + 月["profit"]
        月["combinedSales"] = 月["combinedRevenue"]

    月一覧 = sorted(月表.keys(), reverse=True)
    商品名 = sorted({n for 月 in 月表.values() for n in 月["products"]})

    return {
        "status": "ok",
        "months": 月一覧,
        "products": 商品名,
        "menuTypes": list(メニュー種別),
        "matrix": 月表,
        # **この2つは会員の表から作る。**まだ移していないので空で返す。
        # GAS 側が転送するのは会員を移したあと。
        "registrationRoutes": {},
        "categoryUsage": {},
    }


def _だしか含めて内訳(月, 名, 個数, 単価, 原価, 売上, 粗利):
    """天然だしなら値引きの段ごとに、それ以外は売価ごとに内訳を作る。"""
    if _だしか(名):
        内訳 = だしの内訳(個数)
        for 段 in だしの段:
            数 = 内訳.get(段["key"], 0)
            if not 数:
                continue
            段売上 = 段["price"] * 数
            段原価 = 原価 * 数
            _内訳を足す(月, _だしの鍵(名, 段["key"]), 名, 段["price"], 段["label"],
                        数, 段売上, 段原価, 段売上 - 段原価, 段["sortOrder"])
        return
    鍵 = f"{名}|{_整数(単価)}"
    _内訳を足す(月, 鍵, 名, _整数(単価), f"{_整数(単価):,}円",
                個数, 売上, 原価 * 個数, 粗利)
