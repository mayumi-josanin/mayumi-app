"""商品を、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**手順は [表を移す手順](docs/design/表を移す手順.md)。

## 数えてから書いた

    同名の関数   無し
    表またぎ     5つ（updateRecordStatus / updateNoticeVisibility /
                 deleteNoticeListing / deleteRow / deleteRows）
    保存以外     handleAddProduct / handleUpdateProduct が**通知を送る**
    相乗り       getProducts は商品だけを返す
    別経路       **getProductRevenueMasterMap_ が商品シートを直接読んでいる**
                 （分析画面の原価）。GAS 側で getAdminProducts から作る形に直した

## 原価は仕入の表から引く

商品の表に原価の列は無い。GAS は `getProductCostMap_()` で
仕入の表（SupplierPrice）から**商品名で**引いている。ここも同じにする。
名前で引くのは危ういが、**今そうなっているものを移行では変えない。**
"""

from django.db import transaction
from django.utils import timezone

from apps.content.models import Product
from apps.records.models import SupplierPrice


def _文(v):
    return "" if v is None else str(v)


def _数(v, 既定=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 既定


def _時刻の字(d):
    if not d:
        return ""
    return timezone.localtime(d).strftime("%Y-%m-%dT%H:%M:%S+09:00")


def _公開か(値, 既定=True):
    文 = _文(値).strip()
    return 既定 if not 文 else 文 != "非公開"


def _状態の字(公開):
    return "公開" if 公開 else "非公開"


def _原価の表():
    """GAS の getProductCostMap_ と同じ。**商品名で引く。**

    **数として返す。**データベースは Decimal で持っており、そのまま返すと
    JSON では `"1875.000000"` という**文字列**になる。GAS は
    `Number(costMap[name] || 0)` として数を期待している（2026-09-05）。
    小数が出ないもの（1875.0）は整数にする。GAS の見た目に合わせる。
    """
    出 = {}
    for s in SupplierPrice.objects.all():
        名 = (s.product_name or "").strip()
        if not 名:
            continue
        v = float(s.price or 0)
        出[名] = int(v) if v == int(v) else v
    return 出


def _一件(p, 原価):
    from .views import _画像

    画像 = _画像(p.icon_url)
    説明画像 = _画像(p.description_image_url)
    売切 = (p.sold_out or "").strip() or "販売中"
    return {
        "rowIdx": p.sheet_row,
        "category": p.category or "",
        "name": p.name or "",
        "price": p.price or 0,
        "costPrice": 原価.get((p.name or "").strip(), 0),
        "icon": 画像[0] if 画像 else (p.icon_url or "🌿"),
        "imageUrls": 画像,
        "bg": "c1",
        "status": _状態の字(p.published),
        "description": p.description or "",
        "descriptionImage": 説明画像[0] if 説明画像 else (p.description_image_url or ""),
        "descriptionImageUrls": 説明画像,
        "updatedAt": _時刻の字(p.updated_at),
        # **在庫が数字でないときは null のまま返す。**（views.py と同じ理由）
        "stockQty": p.stock,
        "lowStockThreshold": p.stock_warning or 0,
        "soldOutStatus": 売切,
        "isSoldOut": 売切 == "売切",
        "publishAt": _時刻の字(p.publish_at),
        "noticeStatus": _状態の字(p.notice_listed),
        "noticeDeletedAt": _時刻の字(p.notice_delisted_at),
        "sortOrder": p.sort_order or 0,
    }


def 一覧():
    """GAS の getAdminProducts。**消した分は返さない。**並びはシートの順。"""
    原価 = _原価の表()
    return {"status": "ok",
            "products": [_一件(p, 原価)
                         for p in Product.objects.filter(deleted=False).order_by("sheet_row")]}


def 原価の対応表():
    """GAS の getProductRevenueMasterMap_ と同じ。分析画面が使う。

    **削除済みの商品も含める。**`一覧()` は消した分を除くが、こちらは含める。
    過去の売上の記録は、消したあとの商品も指しているため、
    除くと**その分の原価が0になり、粗利が実際より大きく出る。**

    お知らせの `sortOrder` でも同じ形の取りこぼしをした
    （商品管理に無い商品名で落ちていた）。**消えたものを指す記録がある。**
    """
    原価 = _原価の表()
    出 = {}
    for p in Product.objects.all():
        名 = (p.name or "").strip()
        if not 名:
            continue
        出[名] = {
            "name": 名,
            "price": p.price or 0,
            "costPrice": 原価.get(名, 0),
            "status": _状態の字(p.published),
        }
    return {"status": "ok", "products": 出}


def 足す(d):
    名 = _文(d.get("name")).strip()
    if not 名:
        return {"status": "error", "message": "商品名を入力してください"}
    from .views import _画像

    with transaction.atomic():
        最大 = Product.objects.select_for_update().order_by("-sheet_row").values_list(
            "sheet_row", flat=True).first() or 1
        p = Product.objects.create(
            sheet_row=最大 + 1,
            category=_文(d.get("category")).strip(),
            name=名,
            price=_数(d.get("price")),
            icon_url="\n".join(_画像(d.get("imageUrls") or d.get("icon"))) or _文(d.get("icon")).strip(),
            published=_公開か(d.get("status")),
            description=_文(d.get("description")),
            description_image_url="\n".join(_画像(d.get("descriptionImageUrls")
                                                 or d.get("descriptionImage"))),
            stock=_数(d.get("stockQty"), None) if d.get("stockQty") not in (None, "") else None,
            stock_warning=_数(d.get("lowStockThreshold"), None)
            if d.get("lowStockThreshold") not in (None, "") else None,
            sold_out=_文(d.get("soldOutStatus")).strip(),
            notice_listed=_公開か(d.get("noticeStatus") or d.get("status")),
            publish_at=_日時(d.get("publishAt")),
            updated_at=timezone.now(),
        )
    return {"status": "ok", "rowIdx": p.sheet_row,
            "effectiveStatus": _状態の字(p.published),
            "effectivePublishAt": _時刻の字(p.publish_at)}


def _日時(v):
    from django.utils.dateparse import parse_datetime

    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_datetime(文)
    if d and timezone.is_naive(d):
        d = timezone.make_aware(d)
    return d


def 書き換える(d):
    """**送られてこなかった項目は触らない。**"""
    p = _探す(d)
    if not p:
        return {"status": "error", "message": "商品が見つかりませんでした"}
    from .views import _画像

    if d.get("name"):
        p.name = _文(d["name"]).strip()
    if "category" in d:
        p.category = _文(d.get("category")).strip()
    if "price" in d:
        p.price = _数(d.get("price"))
    if "icon" in d or "imageUrls" in d:
        p.icon_url = "\n".join(_画像(d.get("imageUrls") or d.get("icon"))) or _文(d.get("icon")).strip()
    if d.get("status"):
        p.published = _公開か(d["status"])
    if "description" in d:
        p.description = _文(d.get("description"))
    if "descriptionImage" in d or "descriptionImageUrls" in d:
        p.description_image_url = "\n".join(
            _画像(d.get("descriptionImageUrls") or d.get("descriptionImage")))
    if "stockQty" in d:
        p.stock = _数(d.get("stockQty"))
    if "lowStockThreshold" in d:
        p.stock_warning = _数(d.get("lowStockThreshold"))
    if "soldOutStatus" in d:
        p.sold_out = _文(d.get("soldOutStatus")).strip()
    if d.get("noticeStatus"):
        p.notice_listed = _公開か(d["noticeStatus"])
    if "publishAt" in d:
        p.publish_at = _日時(d.get("publishAt"))
    p.updated_at = timezone.now()
    p.save()
    return {"status": "ok",
            "effectiveStatus": _状態の字(p.published),
            "effectivePublishAt": _時刻の字(p.publish_at)}


def 公開を変える(d):
    p = _探す(d)
    if not p:
        return {"status": "error", "message": "商品が見つかりませんでした"}
    p.published = _公開か(d.get("status"))
    p.updated_at = timezone.now()
    p.save(update_fields=["published", "updated_at", "changed_at"])
    return {"status": "ok"}


def 一覧掲載を変える(d):
    p = _探す(d)
    if not p:
        return {"status": "error", "message": "商品が見つかりませんでした"}
    出す = _公開か(d.get("status"))
    p.notice_listed = 出す
    if 出す:
        p.notice_listed_at = timezone.now()
    else:
        p.notice_delisted_at = timezone.now()
    p.save()
    return {"status": "ok"}


def 一覧から外す(d):
    p = _探す(d)
    if not p:
        return {"status": "error", "message": "商品が見つかりませんでした"}
    p.notice_listed = False
    p.notice_delisted_at = timezone.now()
    p.save(update_fields=["notice_listed", "notice_delisted_at", "changed_at"])
    return {"status": "ok"}


def 消す(d):
    """**印を付けるだけ。**（お知らせ・メニューと同じ）"""
    p = _探す(d)
    if not p:
        return {"status": "error", "message": "商品が見つかりませんでした"}
    p.deleted = True
    p.deleted_at = timezone.now()
    p.delete_reason = _文(d.get("reason")).strip() or "管理画面から削除"
    p.save(update_fields=["deleted", "deleted_at", "delete_reason", "changed_at"])
    return {"status": "ok"}


def まとめて消す(d):
    番号 = d.get("rowIdxs") or d.get("rows") or []
    if not isinstance(番号, list):
        return {"status": "error", "message": "行番号の並びが必要です"}
    数 = sum(1 for r in 番号
             if 消す({"rowIdx": r, "reason": d.get("reason")}).get("status") == "ok")
    return {"status": "ok", "deleted": 数}


def _探す(d):
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        return None
    if 行 <= 1:
        return None
    return Product.objects.filter(sheet_row=行).first()
