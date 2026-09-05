"""カレンダーを、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**手順は [表を移す手順](docs/design/表を移す手順.md)。

## 数えてから書いた

    同名の関数   **2組**（getCalendarEvents / getAdminCalendar）
                 効くのは後の定義（8550行 / 8606行）
    表またぎ     5つ
    保存以外     handleAddCalendar / handleUpdateCalendar が**通知を送る**
    相乗り       無し
    落ちていた列 **カテゴリ70件**（書き出し・モデル・取り込み・読み出しの
                 4か所すべてに無かった）

## 日付は年月日だけ

シートには時刻付きで入っているが、アプリは `raw.split(/[ T]/)[0]` で
**年月日しか取り出していない**（app.js 3303行）。時刻は使われていない。
"""

from django.db import transaction
from django.utils import timezone

from apps.content.models import CalendarEvent


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


def _日付の字(d):
    """GAS の formatMaybeDateTime_ は Date を日時の形で返す。"""
    if not d:
        return ""
    return d.strftime("%Y-%m-%dT00:00:00+09:00")


def _日付(v):
    from django.utils.dateparse import parse_date

    文 = _文(v).strip()
    if not 文:
        return None
    return parse_date(文[:10].replace("/", "-"))


def _日時(v):
    from django.utils.dateparse import parse_datetime

    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_datetime(文)
    if d and timezone.is_naive(d):
        d = timezone.make_aware(d)
    return d


def _公開か(値, 既定=True):
    文 = _文(値).strip()
    return 既定 if not 文 else 文 != "非公開"


def _状態の字(公開):
    return "公開" if 公開 else "非公開"


def _一件(c):
    """GAS の getAdminCalendar（**8606行のほう**）が返す1件と同じ形。"""
    from .views import _画像

    画像 = _画像(c.image_url)
    return {
        "rowIdx": c.sheet_row,
        "date": _日付の字(c.event_on),
        "title": c.title or "",
        "desc": c.detail or "",
        "color": c.color or "",
        "category": c.category or "",
        "publishStatus": _状態の字(c.published),
        "image": 画像[0] if 画像 else "",
        "imageUrls": 画像,
        "updatedAt": _時刻の字(c.updated_at),
        "publishAt": _時刻の字(c.publish_at),
        "linkUrl": c.link_url or "",
        "linkButtonText": c.button_text or "",
        "menuRowIdx": c.menu_row or 0,
        "noticeStatus": _状態の字(c.notice_listed),
        "noticeDeletedAt": _時刻の字(c.notice_delisted_at),
        "sortOrder": c.sort_order or 0,
    }


def 一覧():
    """GAS の getAdminCalendar。**消した分は返さない。**並びはシートの順。"""
    return {"status": "ok",
            "events": [_一件(c)
                       for c in CalendarEvent.objects.filter(deleted=False).order_by("sheet_row")]}


def 足す(d):
    題 = _文(d.get("title")).strip()
    if not 題:
        return {"status": "error", "message": "イベント名を入力してください"}
    from .views import _画像

    with transaction.atomic():
        最大 = CalendarEvent.objects.select_for_update().order_by("-sheet_row").values_list(
            "sheet_row", flat=True).first() or 1
        c = CalendarEvent.objects.create(
            sheet_row=最大 + 1,
            event_on=_日付(d.get("date")) or timezone.localdate(),
            title=題,
            detail=_文(d.get("desc")),
            color=_文(d.get("color")).strip(),
            category=_文(d.get("category")).strip(),
            published=_公開か(d.get("publishStatus") or d.get("status")),
            image_url="\n".join(_画像(d.get("imageUrls") or d.get("image"))),
            link_url=_文(d.get("linkUrl")).strip(),
            button_text=_文(d.get("linkButtonText")).strip(),
            menu_row=_数(d.get("menuRowIdx")) or None,
            notice_listed=_公開か(d.get("noticeStatus") or d.get("publishStatus")),
            publish_at=_日時(d.get("publishAt")),
            updated_at=timezone.now(),
        )
    return {"status": "ok", "rowIdx": c.sheet_row,
            "effectiveStatus": _状態の字(c.published),
            "effectivePublishAt": _時刻の字(c.publish_at)}


def 書き換える(d):
    """**送られてこなかった項目は触らない。**"""
    c = _探す(d)
    if not c:
        return {"status": "error", "message": "イベントが見つかりませんでした"}
    from .views import _画像

    if d.get("date"):
        c.event_on = _日付(d["date"]) or c.event_on
    if d.get("title"):
        c.title = _文(d["title"]).strip()
    if "desc" in d:
        c.detail = _文(d.get("desc"))
    if "color" in d:
        c.color = _文(d.get("color")).strip()
    if "category" in d:
        c.category = _文(d.get("category")).strip()
    if d.get("publishStatus") or d.get("status"):
        c.published = _公開か(d.get("publishStatus") or d.get("status"))
    if "image" in d or "imageUrls" in d:
        c.image_url = "\n".join(_画像(d.get("imageUrls") or d.get("image")))
    if "linkUrl" in d:
        c.link_url = _文(d.get("linkUrl")).strip()
    if "linkButtonText" in d:
        c.button_text = _文(d.get("linkButtonText")).strip()
    if "menuRowIdx" in d:
        c.menu_row = _数(d.get("menuRowIdx")) or None
    if d.get("noticeStatus"):
        c.notice_listed = _公開か(d["noticeStatus"])
    if "publishAt" in d:
        c.publish_at = _日時(d.get("publishAt"))
    c.updated_at = timezone.now()
    c.save()
    return {"status": "ok",
            "effectiveStatus": _状態の字(c.published),
            "effectivePublishAt": _時刻の字(c.publish_at)}


def 公開を変える(d):
    c = _探す(d)
    if not c:
        return {"status": "error", "message": "イベントが見つかりませんでした"}
    c.published = _公開か(d.get("status"))
    c.updated_at = timezone.now()
    c.save(update_fields=["published", "updated_at", "changed_at"])
    return {"status": "ok"}


def 一覧掲載を変える(d):
    c = _探す(d)
    if not c:
        return {"status": "error", "message": "イベントが見つかりませんでした"}
    出す = _公開か(d.get("status"))
    c.notice_listed = 出す
    if 出す:
        c.notice_listed_at = timezone.now()
    else:
        c.notice_delisted_at = timezone.now()
    c.save()
    return {"status": "ok"}


def 一覧から外す(d):
    c = _探す(d)
    if not c:
        return {"status": "error", "message": "イベントが見つかりませんでした"}
    c.notice_listed = False
    c.notice_delisted_at = timezone.now()
    c.save(update_fields=["notice_listed", "notice_delisted_at", "changed_at"])
    return {"status": "ok"}


def 消す(d):
    """**印を付けるだけ。**（お知らせ・メニュー・商品と同じ）"""
    c = _探す(d)
    if not c:
        return {"status": "error", "message": "イベントが見つかりませんでした"}
    c.deleted = True
    c.deleted_at = timezone.now()
    c.delete_reason = _文(d.get("reason")).strip() or "管理画面から削除"
    c.save(update_fields=["deleted", "deleted_at", "delete_reason", "changed_at"])
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
    return CalendarEvent.objects.filter(sheet_row=行).first()
