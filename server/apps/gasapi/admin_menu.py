"""メニュー（ホーム）を、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**手順は [表を移す手順](docs/design/表を移す手順.md)。

## 数えてから書いた

    同名の関数   無し
    表またぎ     handleUpdateNoticeVisibility / handleDeleteNoticeListing の2つ
                 （BLOG / PRODUCTS / CALENDAR / MENUS で使い回されている）
    保存以外     handleAddMenu / handleUpdateMenu が**通知を送る**
    大元         getMenus / getAdminMenus は別々。相乗りしている表は無い

## 並べ替えは「中身の入れ替え」

GAS の `handleMoveMenu` は、**行の中身を丸ごと隣の行と入れ替える。**
行番号は動かない。ここでも同じにする。表示順の列を書き換える形にすると、
管理アプリが持っている行番号との対応が崩れる。

## 削除は印を付けるだけ

GAS は `markRowSoftDeleted_`。お知らせと同じ。（FAQ・カテゴリは本当に消す）
"""

from django.db import transaction
from django.utils import timezone

from apps.content.models import Menu


def _文(v):
    return "" if v is None else str(v)


def _時刻の字(d):
    """**必ず日本時間へ直してから並べる。**（お知らせで9時間ずれた）"""
    if not d:
        return ""
    return timezone.localtime(d).strftime("%Y-%m-%dT%H:%M:%S+09:00")


def _日付の字(d):
    """GAS の formatMaybeDateTime_ は、日付だけの値も日時の形で返す。"""
    if not d:
        return ""
    return d.strftime("%Y-%m-%dT00:00:00+09:00")


def _公開か(値, 既定=True):
    文 = _文(値).strip()
    return 既定 if not 文 else 文 != "非公開"


def _状態の字(公開):
    return "公開" if 公開 else "非公開"


def _画像(d, いま=None):
    """**views.py の _画像 を使う。**自前に書かない（お知らせで82件壊した）。"""
    from .views import _画像 as 解く

    並び = d.get("imageUrls")
    if isinstance(並び, list):
        return [x for x in (解く(並び)) if x]
    if "imageUrl" in d:
        return 解く(d.get("imageUrl"))
    return list(いま or [])


def _一件(m):
    """GAS の getAdminMenus が返す1件と同じ形。**項目名を変えない。**"""
    画像 = list(m.image_urls or [])
    return {
        "rowIdx": m.sheet_row,
        "date": _日付の字(m.registered_on),
        "name": m.name or "",
        "imageUrl": 画像[0] if 画像 else "",
        "imageUrls": 画像,
        "description": m.summary or "",
        "reservationStatus": m.booking_status or "",
        "publishStatus": _状態の字(m.published),
        "category": m.category or "",
        "updatedAt": _時刻の字(m.updated_at),
        "publishAt": _時刻の字(m.publish_at),
        "noticeStatus": _状態の字(m.notice_listed),
        "noticeDeletedAt": _時刻の字(m.notice_delisted_at),
        "sortOrder": m.sort_key or 0,
    }


def 一覧():
    """GAS の getAdminMenus。**消した分は返さない。**並びはシートの順。"""
    return {"status": "ok",
            "menus": [_一件(m) for m in Menu.objects.filter(deleted=False).order_by("sheet_row")]}


def 足す(d):
    """GAS の handleAddMenu。**既定は非公開**（GAS がそうしている）。"""
    名 = _文(d.get("name")).strip()
    if not 名:
        return {"status": "error", "message": "メニュー名を入力してください"}

    with transaction.atomic():
        最大 = Menu.objects.select_for_update().order_by("-sheet_row").values_list(
            "sheet_row", flat=True).first() or 1
        m = Menu.objects.create(
            sheet_row=最大 + 1,
            name=名,
            image_urls=_画像(d),
            summary=_文(d.get("description")),
            booking_status=_文(d.get("reservationStatus")).strip() or "予約対象外",
            published=_公開か(d.get("publishStatus"), 既定=False),
            category=_文(d.get("category")).strip(),
            notice_listed=_公開か(d.get("noticeStatus") or d.get("publishStatus"), 既定=True),
            publish_at=_日時(d.get("publishAt")),
            registered_on=timezone.localdate(),
            updated_at=timezone.now(),
        )
    return {"status": "ok", "rowIdx": m.sheet_row,
            "effectiveStatus": _状態の字(m.published),
            "effectivePublishAt": _時刻の字(m.publish_at)}


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
    """GAS の handleUpdateMenu。**送られてこなかった項目は触らない。**"""
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "メニューが見つかりませんでした"}

    if d.get("name"):
        m.name = _文(d["name"]).strip()
    if "imageUrl" in d or "imageUrls" in d:
        m.image_urls = _画像(d, m.image_urls)
    if "description" in d:
        m.summary = _文(d.get("description"))
    if d.get("reservationStatus"):
        m.booking_status = _文(d["reservationStatus"]).strip()
    if d.get("publishStatus"):
        m.published = _公開か(d["publishStatus"])
    if "category" in d:
        m.category = _文(d.get("category")).strip()
    if d.get("noticeStatus"):
        m.notice_listed = _公開か(d["noticeStatus"])
    if "publishAt" in d:
        m.publish_at = _日時(d.get("publishAt"))
    m.updated_at = timezone.now()
    m.save()
    return {"status": "ok",
            "effectiveStatus": _状態の字(m.published),
            "effectivePublishAt": _時刻の字(m.publish_at)}


def 一覧掲載を変える(d):
    """GAS の handleUpdateNoticeVisibility のうち、メニューのぶん。"""
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "メニューが見つかりませんでした"}
    出す = _公開か(d.get("status"))
    m.notice_listed = 出す
    if 出す:
        m.notice_listed_at = timezone.now()
    else:
        m.notice_delisted_at = timezone.now()
    m.save()
    return {"status": "ok"}


def 一覧から外す(d):
    """GAS の handleDeleteNoticeListing のうち、メニューのぶん。

    **メニューそのものは消さない。**お知らせ一覧から下ろすだけ。
    """
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "メニューが見つかりませんでした"}
    m.notice_listed = False
    m.notice_delisted_at = timezone.now()
    m.save(update_fields=["notice_listed", "notice_delisted_at", "changed_at"])
    return {"status": "ok"}


def 消す(d):
    """GAS の handleDeleteMenu。**印を付けるだけ。**"""
    m = _探す(d)
    if not m:
        return {"status": "error", "message": "削除失敗"}
    m.deleted = True
    m.deleted_at = timezone.now()
    m.delete_reason = "管理画面から削除"
    m.save(update_fields=["deleted", "deleted_at", "delete_reason", "changed_at"])
    return {"status": "ok"}


def 動かす(d):
    """GAS の handleMoveMenu。**行の中身を隣と入れ替える。行番号は動かない。**"""
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        行 = 0
    向き = _文(d.get("direction")).strip()
    if 行 <= 1:
        return {"status": "error", "message": "Invalid row index"}

    先 = 行 - 1 if 向き == "up" else 行 + 1
    if 先 <= 1:
        return {"status": "error", "message": "Cannot move further in this direction"}

    with transaction.atomic():
        a = Menu.objects.select_for_update().filter(sheet_row=行).first()
        b = Menu.objects.select_for_update().filter(sheet_row=先).first()
        if not a or not b:
            return {"status": "error", "message": "Cannot move further in this direction"}

        # **中身だけを入れ替える。**行番号（sheet_row）はそのまま。
        項目 = ["registered_on", "name", "summary", "category", "booking_status",
                "image_urls", "sort_key", "published", "notice_listed", "deleted",
                "deleted_at", "delete_reason", "publish_at", "notice_listed_at",
                "notice_delisted_at", "sort_order", "image_url"]
        for f in 項目:
            va, vb = getattr(a, f), getattr(b, f)
            setattr(a, f, vb)
            setattr(b, f, va)
        いま = timezone.now()
        a.updated_at = いま
        b.updated_at = いま
        a.save()
        b.save()
    return {"status": "ok"}


def _探す(d):
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        return None
    if 行 <= 1:
        return None
    return Menu.objects.filter(sheet_row=行).first()
