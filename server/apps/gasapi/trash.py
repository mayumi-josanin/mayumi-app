"""ゴミ箱（消した印の付いた行）。GAS の getAdminTrashItems / handleRestoreDeletedRecord / handleHardDeleteRecord。

GAS は「削除状態」「削除日時」の列で消した行を見分けている。サーバーでは各表の
`deleted` / `deleted_at` にあたる。**種類の名前と並び（会員・注文・NEWS・ショップ・
カレンダー・ホーム・Push通知）は GAS の ADMIN_TRASH_SOURCES と同じ。**

## 注文はここに来ない

GAS の handleDeleteOrders は行を本当に消している（orders.注文を消す も同じ）。
消した印が無いので、ゴミ箱には注文が並ばない。種類の選択肢に「注文」を残しているのは
旧管理アプリと同じ並びにするため。

## 会員は行番号ではなく会員IDで指す

GAS はシートの行番号（rowIdx）で指すが、サーバーの会員は会員IDが主キー。
会員だけ `memberId` で受ける（rowIdx が来ても会員IDとして読む）。
"""

from django.db.models import Q
from django.utils import timezone

from apps.content.models import CalendarEvent, Menu, News, Product, PushNotice
from apps.members.models import Member

# GAS の ADMIN_TRASH_SOURCES と同じ名前。並びは getAdminTrashItems と同じ。
種類の名前 = {
    "USERS": "会員",
    "ORDERS": "注文",
    "BLOG": "NEWS",
    "PRODUCTS": "ショップ",
    "CALENDAR": "カレンダー",
    "MENUS": "ホーム",
    "PUSH": "Push通知",
}

# (表, 題名の欄, 題名が空のときの名前) … GAS の getTrashRowLabel_ と同じ
_表 = {
    "BLOG": (News, "title", "NEWS"),
    "PRODUCTS": (Product, "name", "商品"),
    "CALENDAR": (CalendarEvent, "title", "イベント"),
    "MENUS": (Menu, "name", "メニュー"),
    "PUSH": (PushNotice, "title", "通知"),
}


def _文(v):
    return "" if v is None else str(v)


def _時刻の字(d):
    return timezone.localtime(d).strftime("%Y-%m-%dT%H:%M:%S+09:00") if d else ""


def _消した印(表):
    """印だけ付いて日時が無い行も、日時だけある行も、どちらも「消した」扱い（GAS の isSoftDeletedByColumns_）。"""
    return 表.objects.filter(Q(deleted=True) | Q(deleted_at__isnull=False))


def 一覧():
    """GAS の getAdminTrashItems。消した日時の新しい順。"""
    出 = []
    for m in _消した印(Member):
        出.append({
            "sheet": "USERS", "source": 種類の名前["USERS"],
            "rowIdx": 0, "memberId": m.member_id,
            "title": m.name or m.member_id or "会員",
            "deletedAt": _時刻の字(m.deleted_at),
            # 会員の表に削除理由の欄は無い。統合で消えた方は行き先だけ分かるようにする。
            "reason": f"{m.merged_into_id} へ統合" if _文(m.merged_into_id).strip() else "",
        })
    for 鍵, (表, 欄, 既定) in _表.items():
        for r in _消した印(表):
            出.append({
                "sheet": 鍵, "source": 種類の名前[鍵],
                "rowIdx": r.sheet_row or 0, "memberId": "",
                "title": _文(getattr(r, 欄, "")).strip() or 既定,
                "deletedAt": _時刻の字(r.deleted_at),
                "reason": _文(getattr(r, "delete_reason", "")),
            })
    出.sort(key=lambda x: x["deletedAt"], reverse=True)
    return {"status": "ok", "items": 出}


def _探す(d):
    """sheet と rowIdx（会員は memberId）で1行を指す。無ければ None。"""
    d = d or {}
    鍵 = _文(d.get("sheet")).strip().upper()
    if 鍵 == "USERS":
        会員ID = _文(d.get("memberId") or d.get("rowIdx")).strip()
        return Member.objects.filter(pk=会員ID).first() if 会員ID else None
    組 = _表.get(鍵)
    if not 組:
        return None
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        行 = 0
    return 組[0].objects.filter(sheet_row=行).first() if 行 > 1 else None


def 戻す(d):
    """GAS の handleRestoreDeletedRecord（clearRowSoftDeleted_）。印と日時を外す。会員は統合先も外す。"""
    r = _探す(d)
    if r is None:
        return {"status": "error", "message": "復元対象が見つかりません"}
    r.deleted = False
    r.deleted_at = None
    欄 = ["deleted", "deleted_at"]
    if hasattr(r, "delete_reason"):
        r.delete_reason = ""
        欄.append("delete_reason")
    if hasattr(r, "merged_into_id"):
        r.merged_into_id = ""
        欄.append("merged_into_id")
    if hasattr(r, "changed_at"):
        欄.append("changed_at")
    r.save(update_fields=欄)
    return {"status": "ok"}


def 完全に消す(d):
    """GAS の handleHardDeleteRecord（sheet.deleteRow）。**本当に消す。戻せない。**"""
    r = _探す(d)
    if r is None:
        return {"status": "error", "message": "削除対象が見つかりません"}
    r.delete()
    return {"status": "ok"}
