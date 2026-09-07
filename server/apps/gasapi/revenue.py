"""売上の記録（メニュー収益・商品収益）。**GAS が転送してくる先。**

## なぜ作ることになったか

分析画面（`getAnalyticsData`）だけを先に移して、**その元になる
売上の記録の書き込みを移し忘れた**（2026-09-07）。

    売上の入力  → シート（移していない）
    分析画面    → サーバー（移した）

院長が収益を入力しても、分析に出てこなかった。**入力したものは無事**
だったが、見に行く先が違っていた。

手順書に「**その応答に、他の表が相乗りしていないか**」と自分で書いて
おきながら、注文（0件）ばかり見て**同じ関数が読んでいる売上の表**を
数え落とした。切り替え後の確認も「数字が一致するか」だけで、
**入力してみるところまで試していなかった。**

## 削除は本当に消し、行番号が繰り上がる

GAS は `sheet.deleteRow(rowIdx)`。FAQ と同じ作法。
お知らせ・商品のような論理削除ではない。**表ごとに今の振る舞いに合わせる。**
"""

import re
import unicodedata
from decimal import Decimal

from django.db import transaction
from django.db.models import F
from django.utils import timezone

from apps.records.models import RevenueRecord

メニュー種別 = ["母乳外来", "ビジリス", "教室", "その他"]


def _文(v):
    return "" if v is None else str(v)


def _数(v, 既定=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 既定


def _金(v):
    """**数として返す。**Decimal のまま返すと JSON では文字列になる。"""
    try:
        d = Decimal(str(v or 0))
    except Exception:
        return 0
    return int(d) if d == d.to_integral_value() else float(d)


def _日付の字(d):
    return d.strftime("%Y-%m-%d") if d else ""


def _日付(v):
    from django.utils.dateparse import parse_date

    文 = _文(v).strip()
    if not 文:
        return None
    return parse_date(文[:10].replace("/", "-"))


def _種別をそろえる(v):
    """GAS の normalizeMenuRevenueType_ と同じ。**当てはまらなければ「その他」。**"""
    t = re.sub(r"[\s　]+", "", unicodedata.normalize("NFKC", _文(v)))
    return t if t in メニュー種別 else "その他"


def _一件(r):
    """GAS の getMenuRevenueRecords / getProductRevenueRecords と同じ形。"""
    数 = _数(r.quantity, 1)
    単価 = _金(r.unit_price)
    原価 = _金(r.unit_cost)
    売上 = 単価 * 数
    原価計 = 原価 * 数
    共通 = {
        "rowIdx": r.sheet_row,
        "date": _日付の字(r.recorded_on),
        "unitPrice": 単価,
        "unitCost": 原価,
        "totalAmount": 売上,
        "totalCost": 原価計,
        "profit": 売上 - 原価計,
        "note": r.memo or "",
    }
    if r.kind == RevenueRecord.MENU:
        共通.update({"menuType": r.name or "", "count": 数})
    else:
        共通.update({"productName": r.name or "", "qty": 数})
    return 共通


def _一覧(種別):
    件 = RevenueRecord.objects.filter(kind=種別, deleted=False).order_by("sheet_row")
    return {"status": "ok", "records": [_一件(r) for r in 件]}


def メニューの記録():
    return _一覧(RevenueRecord.MENU)


def 商品の記録():
    return _一覧(RevenueRecord.PRODUCT)


def _保存(種別, d, 名の鍵, 数の鍵):
    """GAS の handleSave…RevenueRecord。**複数まとめて来ることがある。**

    `rowIdx` があれば更新、無ければ追加。管理画面は入力行を並べて
    まとめて送ってくる。
    """
    並び = d.get("records")
    if not isinstance(並び, list) or not 並び:
        並び = [d]

    更新 = 追加 = 0
    with transaction.atomic():
        最大 = RevenueRecord.objects.select_for_update().filter(kind=種別).order_by(
            "-sheet_row").values_list("sheet_row", flat=True).first() or 1
        for item in 並び:
            日 = _日付(item.get("date") or d.get("date"))
            名 = (_種別をそろえる(item.get(名の鍵)) if 種別 == RevenueRecord.MENU
                  else _文(item.get(名の鍵)).strip())
            数 = max(1, _数(item.get(数の鍵), 1))
            単価 = max(0, _数(item.get("unitPrice")))
            原価 = max(0, _数(item.get("unitCost")))
            メモ = _文(item.get("note")).strip()

            行 = item.get("rowIdx", d.get("rowIdx"))
            行 = _数(行, 0)
            if 行 >= 2:
                r = RevenueRecord.objects.select_for_update().filter(
                    kind=種別, sheet_row=行).first()
                if r:
                    r.recorded_on, r.name, r.quantity = 日, 名, 数
                    r.unit_price, r.unit_cost, r.memo = 単価, 原価, メモ
                    r.save()
                    更新 += 1
                    continue
            最大 += 1
            RevenueRecord.objects.create(
                kind=種別, sheet_row=最大, recorded_on=日, name=名,
                quantity=数, unit_price=単価, unit_cost=原価, memo=メモ)
            追加 += 1
    return {"status": "ok", "savedCount": len(並び),
            "updatedCount": 更新, "createdCount": 追加}


def メニューを保存(d):
    return _保存(RevenueRecord.MENU, d, "menuType", "count")


def 商品を保存(d):
    return _保存(RevenueRecord.PRODUCT, d, "productName", "qty")


def _消す(種別, d):
    """GAS の handleDelete…RevenueRecord。**本当に消し、行番号を繰り上げる。**

    シートの deleteRow と同じ。繰り上げないと、管理画面が次に取り直した
    ときの番号とずれていく。（FAQ と同じ作法）
    """
    行 = _数(d.get("rowIdx"), 0)
    if 行 < 2:
        return {"status": "error", "message": "削除対象が不正です"}
    with transaction.atomic():
        r = RevenueRecord.objects.select_for_update().filter(
            kind=種別, sheet_row=行).first()
        if not r:
            return {"status": "error", "message": "削除対象が見つかりません"}
        r.delete()
        # **番号の小さいほうから詰める。**まとめて -1 すると途中で衝突する。
        for x in RevenueRecord.objects.filter(
                kind=種別, sheet_row__gt=行).order_by("sheet_row"):
            RevenueRecord.objects.filter(pk=x.pk).update(sheet_row=F("sheet_row") - 1)
    return {"status": "ok"}


def メニューを消す(d):
    return _消す(RevenueRecord.MENU, d)


def 商品を消す(d):
    return _消す(RevenueRecord.PRODUCT, d)
