"""注文。**GAS が転送してくる先。**

計画は [注文の表を移す](docs/design/注文の表を移す.md)。

## 1つの注文が、商品ごとに複数行

シートがそうなっている。読むときは注文IDでまとめて返す。
**合計金額・支払方法・ステータス・管理メモは最初の1行にだけ入る。**

## 数式を、書き込み時の計算に置き換える

    仕入値  仕入の表（SupplierPrice）から**商品名で引く**
    小計    個数 × 単価
    純利益  個数 × (単価 − 仕入値)

**この置き換えを間違えると、粗利の数字が静かにずれる。**
"""

import re
from decimal import Decimal

from django.db import transaction
from django.utils import timezone

from apps.records.models import OrderLine, SupplierPrice

受付中 = "受付中"
キャンセル = "キャンセル"


def _文(v):
    return "" if v is None else str(v)


def _数(v, 既定=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 既定


def _金(v):
    """数として扱う。**Decimal のまま返すと JSON では文字列になる。**"""
    try:
        d = Decimal(str(v or 0))
    except Exception:
        return 0
    return int(d) if d == d.to_integral_value() else float(d)


def _日時の字(d):
    """GAS の `yyyy/M/d H:mm` と同じ形。"""
    if not d:
        return ""
    t = timezone.localtime(d)
    return f"{t.year}/{t.month}/{t.day} {t.hour}:{t.minute:02d}"


def _原価の表():
    出 = {}
    for s in SupplierPrice.objects.all():
        名 = (s.product_name or "").strip()
        if 名:
            出[名] = Decimal(str(s.price or 0))
    return 出


def _まとめる(行, 管理向け):
    """注文IDごとにまとめる。**並びはシートの順（古い順）。**"""
    表 = {}
    順 = []
    for r in 行:
        if not (r.order_id or "").strip():
            continue
        if r.order_id not in 表:
            順.append(r.order_id)
            共通 = {
                "items": [],
                "total": r.total_label or "",
                "payment": r.payment or "",
                "status": (r.status or 受付中),
                "checked": bool(r.received),
            }
            if 管理向け:
                共通.update({
                    "orderId": r.order_id,
                    "date": _日時の字(r.ordered_at),
                    "customerName": r.customer_name or "",
                    "memberId": r.member_id or "",
                    "internalNote": r.internal_note or "",
                    "rowIndices": [],
                })
            else:
                共通.update({"id": r.order_id, "time": _日時の字(r.ordered_at)})
            表[r.order_id] = 共通
        表[r.order_id]["items"].append({
            "name": r.product_name or "",
            "qty": r.quantity or 0,
            "price": _金(r.unit_price),
        })
        if 管理向け and r.sheet_row:
            表[r.order_id]["rowIndices"].append(r.sheet_row)
    return [表[k] for k in 順]


def お客様の注文(request):
    """GAS の getCustomerOrders。**受け取り済みとキャンセルは出さない。**"""
    from .views import _引数

    会員ID = _文(_引数(request).get("memberId")).strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが必要です"}
    行 = OrderLine.objects.filter(member_id=会員ID).order_by("id")
    生きている = [r for r in 行 if not r.received and (r.status or 受付中) != キャンセル]
    return {"status": "ok", "orders": _まとめる(生きている, 管理向け=False)}


def 管理の注文(request=None):
    """GAS の getAdminOrders。既定では受け取り済みを出さない。"""
    from .views import _引数

    すべて = False
    if request is not None:
        すべて = _引数(request).get("showAll") is True
    行 = OrderLine.objects.all().order_by("id")
    if not すべて:
        行 = [r for r in 行 if not r.received]
    return {"status": "ok", "orders": _まとめる(行, 管理向け=True)}


def 会員の注文(request):
    """GAS の getAdminUserOrders。**その方の全部**（受け取り済みも含む）。"""
    from .views import _引数

    会員ID = _文(_引数(request).get("memberId")).strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが必要です"}
    行 = OrderLine.objects.filter(member_id=会員ID).order_by("id")
    return {"status": "ok", "orders": _まとめる(行, 管理向け=True)}


# ── 書く ───────────────────────────────────────────────

def 注文する(d):
    """GAS の handleOrder。**商品ごとに1行ずつ作る。**

    合計金額・支払方法・ステータス・管理メモは**最初の1行にだけ**入れる。
    シートがそうなっているので、そのまま写す。
    """
    品 = d.get("items")
    if not isinstance(品, list) or not 品:
        return {"status": "error", "message": "ご注文の商品がありません。"}

    注文ID = _文(d.get("orderId")).strip() or ("ORD-" + timezone.now().strftime("%Y%m%d%H%M%S"))
    原価 = _原価の表()
    いま = timezone.now()

    with transaction.atomic():
        if OrderLine.objects.filter(order_id=注文ID).exists():
            # **二重に受け取らない。**通信が失敗して再送されることがある。
            return {"status": "ok", "orderId": 注文ID, "duplicated": True}

        作った = []
        for i, item in enumerate(品):
            名 = _文(item.get("name")).strip()
            個数 = _数(item.get("qty"))
            単価 = Decimal(str(item.get("price") or 0))
            仕入 = 原価.get(名, Decimal("0"))
            作った.append(OrderLine(
                order_id=注文ID,
                ordered_at=いま,
                customer_name=_文(d.get("customerName")).strip(),
                product_name=名,
                quantity=個数,
                unit_price=単価,
                # **数式ではなく、ここで計算する**
                cost_price=仕入,
                subtotal=単価 * 個数,
                profit=(単価 - 仕入) * 個数,
                total_label=("¥{:,}".format(_数(d.get("total"))) if i == 0 else ""),
                payment=(_文(d.get("payment")) if i == 0 else ""),
                status=(受付中 if i == 0 else ""),
                received=False,
                internal_note=(_文(d.get("internalNote")) if i == 0 else ""),
                member_id=_文(d.get("memberId")).strip(),
            ))
        OrderLine.objects.bulk_create(作った)
    return {"status": "ok", "orderId": 注文ID}


def 取り消す(d):
    """GAS の handleCancel。**印を付けるだけ。行は残す。**"""
    注文ID = _文(d.get("orderId")).strip()
    if not 注文ID:
        return {"status": "error", "message": "注文IDが必要です"}
    with transaction.atomic():
        行 = list(OrderLine.objects.select_for_update().filter(order_id=注文ID).order_by("id"))
        if not 行:
            return {"status": "error", "message": "ご注文が見つかりませんでした。"}
        行[0].status = キャンセル
        行[0].save(update_fields=["status", "changed_at"])
    return {"status": "ok"}


def 受け取りを報告する(d):
    """GAS の handleConfirmReceipt。お客様が「受け取りました」と押したとき。"""
    注文ID = _文(d.get("orderId")).strip()
    if not 注文ID:
        return {"status": "error", "message": "注文IDが必要です"}
    with transaction.atomic():
        行 = list(OrderLine.objects.select_for_update().filter(order_id=注文ID))
        if not 行:
            return {"status": "error", "message": "ご注文が見つかりませんでした。"}
        for r in 行:
            r.received = True
            r.save(update_fields=["received", "changed_at"])
    return {"status": "ok"}


def 注文を書き換える(d):
    """GAS の handleUpdateOrder / updateAdminOrder。受付が直すとき。"""
    注文ID = _文(d.get("orderId")).strip()
    if not 注文ID:
        return {"status": "error", "message": "注文IDが必要です"}
    with transaction.atomic():
        行 = list(OrderLine.objects.select_for_update().filter(order_id=注文ID).order_by("id"))
        if not 行:
            return {"status": "error", "message": "ご注文が見つかりませんでした。"}
        頭 = 行[0]
        if "status" in d:
            頭.status = _文(d.get("status")).strip() or 受付中
        if "payment" in d:
            頭.payment = _文(d.get("payment")).strip()
        if "internalNote" in d:
            頭.internal_note = _文(d.get("internalNote"))
        if "checked" in d:
            受け取り = bool(d.get("checked"))
            for r in 行:
                r.received = 受け取り
                r.save(update_fields=["received", "changed_at"])
        頭.save()
    return {"status": "ok"}


def 注文を消す(d):
    """GAS の handleDeleteOrders。**本当に消す。**シートも deleteRow している。"""
    並び = d.get("orderIds") or ([d.get("orderId")] if d.get("orderId") else [])
    並び = [_文(x).strip() for x in 並び if _文(x).strip()]
    if not 並び:
        return {"status": "error", "message": "注文IDが必要です"}
    with transaction.atomic():
        数 = OrderLine.objects.filter(order_id__in=並び).delete()[0]
    return {"status": "ok", "deleted": 数}


def 会員ごとの集計():
    """GAS の buildOrderStatsByMemberId_。会員一覧に出す注文の数。"""
    出 = {}
    for r in OrderLine.objects.exclude(member_id=""):
        s = 出.setdefault(r.member_id, {"orderCount": 0, "pendingOrderCount": 0,
                                        "lastOrderAt": "", "orderTotal": 0, "_見た": set()})
        if r.order_id in s["_見た"]:
            continue
        s["_見た"].add(r.order_id)
        s["orderCount"] += 1
        if not r.received and (r.status or 受付中) != キャンセル:
            s["pendingOrderCount"] += 1
        字 = _日時の字(r.ordered_at)
        if 字 > s["lastOrderAt"]:
            s["lastOrderAt"] = 字
    for s in 出.values():
        s.pop("_見た", None)
    # 合計金額は行ごとの小計から
    for mid in 出:
        合 = sum(_金(r.subtotal) for r in OrderLine.objects.filter(member_id=mid))
        出[mid]["orderTotal"] = 合
    return 出
