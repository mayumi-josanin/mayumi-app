"""カレンダー管理（お客様アプリのカレンダーに出る予定・休診など）。

読み書きは gasapi/admin_calendar.py（GAS の転送先と同じ）。
追加は **日付を複数選べる**（選んだ数だけ行を作る。GAS の handleAddCalendar と同じ）。
"""

from datetime import date

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import CalendarEvent, Menu
from apps.gasapi import admin_calendar
from apps.gasapi.views import _画像

from .permissions import owner_required
from . import images, push

# よく使う色（旧管理アプリの説明文にあったもの）
COLORS = [("#e57373", "休診（赤系）"), ("#f48fb1", "教室（ピンク系）"), ("#81c784", "イベント（緑系）"),
          ("#64b5f6", "青系"), ("#ffb74d", "橙系"), ("#9575cd", "紫系")]


def _候補の分類():
    名 = set(CalendarEvent.objects.filter(deleted=False).exclude(category="").values_list("category", flat=True))
    return sorted(n for n in 名 if n)


def _メニュー候補():
    return list(Menu.objects.filter(deleted=False).order_by("sheet_row").values_list("sheet_row", "name"))


@owner_required
def calendar_list(request):
    today = timezone.localdate()
    try:
        year = int(request.GET.get("year") or today.year)
        month = int(request.GET.get("month") or 0)
    except ValueError:
        year, month = today.year, 0
    rows = CalendarEvent.objects.filter(deleted=False, event_on__year=year).order_by("event_on", "sheet_row")
    if month:
        rows = rows.filter(event_on__month=month)
    menus = dict(_メニュー候補())
    items = [{
        "c": c, "image": (_画像(c.image_url) or [""])[0],
        "menu": menus.get(c.menu_row or 0, ""),
        "past": c.event_on and c.event_on < today,
    } for c in rows]
    years = sorted({d.year for d in CalendarEvent.objects.filter(deleted=False).exclude(event_on=None)
                    .values_list("event_on", flat=True)} | {today.year, today.year + 1})
    return render(request, "manage/calendar_list.html", {
        "items": items, "year": year, "month": month, "years": years, "months": range(1, 13),
        "today": today,
    })


def _入力(request, c: CalendarEvent | None):
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "calendar") for f in request.FILES.getlist("images")]
    status = request.POST.get("status") or "公開"
    dates = [d for d in request.POST.getlist("dates") if d.strip()]
    return {
        "rowIdx": c.sheet_row if c else 0,
        "date": request.POST.get("date", "").strip(),
        "dates": [] if c else dates,
        "title": request.POST.get("title", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "menuRowIdx": request.POST.get("menuRowIdx") or 0,
        "desc": request.POST.get("desc", ""),
        "color": request.POST.get("color", "").strip(),
        "imageUrls": kept + added,
        "image": (kept + added)[0] if (kept + added) else "",
        "status": status,
        "publishStatus": status,
        "noticeStatus": "公開" if request.POST.get("notice_listed") else "非公開",
        "publishAt": request.POST.get("publishAt", ""),
        "linkUrl": request.POST.get("linkUrl", "").strip(),
        "linkButtonText": request.POST.get("linkButtonText", "").strip(),
    }


def _通知(request, 答, d):
    """GAS の _カレンダーの通知を送る_ と同じ条件。"""
    if not request.POST.get("send_push"):
        return
    状態 = 答.get("effectiveStatus") or d.get("status") or "公開"
    if 状態 == "非公開":
        return
    公開日時 = 答.get("effectivePublishAt") or ""
    if 公開日時:
        from apps.gasapi.admin_news import _日時

        t = _日時(公開日時)
        if t and t > timezone.now():
            return
    if push.全員へ送る("📅 " + (d.get("title") or "カレンダー更新"), "カレンダーが更新されました", page="calendar"):
        messages.info(request, "お客様のアプリへ通知を送りました。")
    else:
        messages.error(request, "通知を送れませんでした（予定は保存されています）。")


def _画面(request, c, values, existing):
    return render(request, "manage/calendar_form.html", {
        "event": c, "values": values, "existing_images": existing,
        "categories": _候補の分類(), "menus": _メニュー候補(), "colors": COLORS,
    })


@owner_required
def calendar_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        if not d["dates"] and not d["date"]:
            messages.error(request, "日付を少なくとも1つ選んでください。")
            return _画面(request, None, request.POST, [])
        答 = admin_calendar.足す(d)
        if 答.get("status") == "ok":
            n = 答.get("created") or 1
            messages.success(request, f"予定を{n}件追加しました。" if d["status"] == "公開" else f"下書きとして{n}件保存しました。")
            _通知(request, 答, d)
            return redirect("manage:calendar_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
        return _画面(request, None, request.POST, [])
    return _画面(request, None, {"color": "#e57373", "notice_listed": "on", "date": timezone.localdate().isoformat()}, [])


@owner_required
def calendar_edit(request, row: int):
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, c)
        答 = admin_calendar.書き換える(d)
        if 答.get("status") == "ok":
            messages.success(request, "予定を更新しました。")
            _通知(request, 答, d)
            return redirect(f"/manage/calendar/?year={c.event_on.year}&month={c.event_on.month}" if c.event_on else "/manage/calendar/")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    values = request.POST or {
        "date": c.event_on.isoformat() if c.event_on else "", "title": c.title, "category": c.category,
        "menuRowIdx": c.menu_row or "", "desc": c.detail, "color": c.color or "#e57373",
        "publishAt": timezone.localtime(c.publish_at).strftime("%Y-%m-%dT%H:%M") if c.publish_at else "",
        "linkUrl": c.link_url, "linkButtonText": c.button_text,
        "notice_listed": "on" if c.notice_listed else "",
    }
    return _画面(request, c, values, _画像(c.image_url))


def _戻る(c):
    if c.event_on:
        return f"/manage/calendar/?year={c.event_on.year}&month={c.event_on.month}"
    return "/manage/calendar/"


@owner_required
@require_POST
def calendar_toggle(request, row: int):
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    答 = admin_calendar.公開を変える({"rowIdx": row, "status": "非公開" if c.published else "公開"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{c.title}」を{'非公開' if c.published else '公開'}にしました。")
    return redirect(_戻る(c))


@owner_required
@require_POST
def calendar_delete(request, row: int):
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    答 = admin_calendar.消す({"rowIdx": row, "reason": f"管理画面から（{request.user.username}）"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{c.title}」を削除しました。")
    return redirect(_戻る(c))
