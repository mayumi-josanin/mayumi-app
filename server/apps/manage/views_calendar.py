"""カレンダー管理。**中身は旧管理アプリ（#page-calendar カレンダー管理）と同じ。**

旧アプリで確認したもの: loadAdminCalendar / applyCalendarFilters / renderCalendarList /
renderMonthCalendar / openDayDetailModal / submitCalendar / editCalendar /
refreshCalendarMenuOptions / updateRecentColorsHistory / updateCalendarStatus /
deleteCalendar / bulkDelete('CALENDAR') / previewCalendarDraft。

読み書きは gasapi/admin_calendar.py（GAS の転送先と同じ）。
追加は **日付を複数選べる**（選んだ数だけ行を作る。GAS の handleAddCalendar と同じ）。
"""

import re

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import CalendarEvent, Menu
from apps.gasapi import admin_calendar
from apps.gasapi.views import _画像

from . import images, push
from .permissions import owner_required

# 旧アプリの既定色（休診日の赤系）
DEFAULT_COLOR = "#e57373"


def _区分(c: CalendarEvent) -> str:
    """区分の見分け方。**旧アプリ（isAdminNoticeCalendar〜Event）と同じ規則。**

    区分（カテゴリ）が入っていればそれに従い、空のときだけ題と説明から当てる。
    """
    分類 = (c.category or "").strip()
    題 = c.title or ""
    if 分類:
        if re.search(r"休診|休み|休業", 分類):
            return "holiday"
        if "訪問産後ケア" in 分類:
            return "postpartum"
        if "往診" in 分類:
            return "visit"
        return "event"
    if re.search(r"休診|休み|休業", 題 + " " + (c.detail or "")):
        return "holiday"
    if "訪問産後ケア" in 題:
        return "postpartum"
    if "往診" in 題:
        return "visit"
    return "event"


def _平文(desc: str) -> str:
    """詳細をプレーンにする（旧アプリの stripCalendarDescPlain）。"""
    if not desc:
        return ""
    t = re.sub(r"<br\s*/?>", " ", str(desc), flags=re.I)
    t = re.sub(r"</(p|div|li)>", " ", t, flags=re.I)
    t = re.sub(r"<[^>]+>", "", t)
    t = (t.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<")
         .replace("&gt;", ">").replace("&quot;", '"').replace("&#39;", "'"))
    return re.sub(r"\s+", " ", t).strip()


def _抜粋(desc: str, 最大=60) -> str:
    """一覧用の切り詰め（旧アプリの buildCalendarDescSnippet）。空なら「―」。"""
    平 = _平文(desc)
    if not 平:
        return "―"
    return 平 if len(平) <= 最大 else 平[:最大] + "…"


def _一件(c: CalendarEvent) -> dict:
    画像 = _画像(c.image_url)
    return {
        "c": c,
        "image": 画像[0] if 画像 else "",
        "image_urls": 画像,
        "snippet": _抜粋(c.detail),
        "plain": _平文(c.detail),
        "kind": _区分(c),
        "date": c.event_on.isoformat() if c.event_on else "",
        "color": c.color or DEFAULT_COLOR,
    }


def _最近の色():
    """全イベントから使った色を重複なく（旧アプリの updateRecentColorsHistory）。"""
    見た = []
    for 色 in CalendarEvent.objects.filter(deleted=False).order_by("sheet_row").values_list("color", flat=True):
        色 = (色 or "").strip().lower()
        if 色 and re.fullmatch(r"#[0-9a-f]{6}", 色) and 色 not in 見た:
            見た.append(色)
    return 見た


def _メニュー候補(current: int = 0):
    """「対象メニュー」の選択肢。**開催予定を出すのはイベントだけ**なので、
    カテゴリが「イベント」のメニューだけを出す。既に結びついているものは
    外して選び直せるよう残す（旧アプリの refreshCalendarMenuOptions）。"""
    候補 = []
    for m in Menu.objects.filter(deleted=False).order_by("sheet_row"):
        if (m.category or "").strip() != "イベント" and m.sheet_row != current:
            continue
        候補.append((m.sheet_row, (m.name or "").strip() or f"（無題 行{m.sheet_row}）"))
    return 候補


def _年月(request, today):
    """絞り込み。年は既定で今年（旧アプリの初期値と同じ）、月は既定で全月。"""
    年 = request.GET.get("year") or str(today.year)
    月 = request.GET.get("month") or "all"
    if 年 != "all":
        try:
            年 = int(年)
        except ValueError:
            年 = today.year
    if 月 != "all":
        try:
            月 = int(月)
            if not 1 <= 月 <= 12:
                月 = "all"
        except ValueError:
            月 = "all"
    return 年, 月


@owner_required
def calendar_list(request):
    today = timezone.localdate()
    year, month = _年月(request, today)
    all_items = [_一件(c) for c in CalendarEvent.objects.filter(deleted=False)]
    # 日付順（降順）。日付の無いものは絞り込みの対象外で、末尾に出す。
    all_items.sort(key=lambda it: (it["date"], it["c"].sheet_row), reverse=True)
    shown = []
    for it in all_items:
        d = it["c"].event_on
        if d:
            if year != "all" and d.year != year:
                continue
            if month != "all" and d.month != month:
                continue
        shown.append(it)
    drafts = [it for it in shown if not it["c"].published]
    published = [it for it in shown if it["c"].published]

    years = sorted({it["c"].event_on.year for it in all_items if it["c"].event_on}
                   | {today.year - 1, today.year, today.year + 1})
    # 月間カレンダーの初期表示。絞り込みで年・月が決まっていればそれに従う。
    cal_year = year if year != "all" else today.year
    cal_month = month if month != "all" else today.month
    events_json = [{
        "rowIdx": it["c"].sheet_row, "date": it["date"], "title": it["c"].title or "",
        "category": it["c"].category or "", "desc": it["c"].detail or "", "color": it["color"],
        "kind": it["kind"], "status": "公開" if it["c"].published else "非公開",
        "imageUrls": it["image_urls"], "editUrl": f"/manage/calendar/{it['c'].sheet_row}/",
    } for it in all_items]
    return render(request, "manage/calendar_list.html", {
        "drafts": drafts, "published": published, "year": year, "month": month, "years": years,
        "months": range(1, 13), "today": today, "cal_year": cal_year, "cal_month": cal_month,
        "events_json": events_json,
        "meta": (f"{len(shown)} / {len(all_items)} 件を表示" if all_items else "表示条件に一致するイベントはありません"),
    })


def _入力(request, c: CalendarEvent | None):
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "calendar") for f in request.FILES.getlist("images")]
    status = request.POST.get("status") or "公開"
    dates = [d for d in request.POST.getlist("dates") if d.strip()]
    d = {
        "rowIdx": c.sheet_row if c else 0,
        "date": request.POST.get("date", "").strip(),
        "dates": [] if c else dates,
        "title": request.POST.get("title", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "menuRowIdx": request.POST.get("menuRowIdx") or 0,
        "desc": request.POST.get("desc", ""),
        "color": request.POST.get("color", "").strip() or DEFAULT_COLOR,
        "imageUrls": kept + added,
        "image": (kept + added)[0] if (kept + added) else "",
        "status": status,
        "publishStatus": status,
        "publishAt": request.POST.get("publishAt", ""),
        "linkUrl": request.POST.get("linkUrl", "").strip(),
        "linkButtonText": request.POST.get("linkButtonText", "").strip(),
    }
    # 「この保存をお知らせ一覧に反映する」。
    # 新規: チェックなら公開設定と同じ、外せばお知らせ一覧には出さない。
    # 編集: チェックのときだけ触る（外していればイベント情報だけ更新し、掲載は変えない）。
    反映 = bool(request.POST.get("notice_refresh"))
    if not c:
        d["noticeStatus"] = status if 反映 else "非公開"
    elif 反映:
        d["noticeStatus"] = status
    return d


def _通知(request, 答, d):
    """GAS の _カレンダーの通知を送る_ と同じ条件（公開保存のときだけ）。"""
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
    current = 0
    try:
        current = int(values.get("menuRowIdx") or 0)
    except (TypeError, ValueError):
        pass
    # 追加中に弾かれて画面に戻るとき、足してあった日付を消さないため。
    picked = values.getlist("dates") if hasattr(values, "getlist") else list(values.get("dates") or [])
    return render(request, "manage/calendar_form.html", {
        "event": c, "values": values, "existing_images": existing, "picked_dates": picked,
        "menus": _メニュー候補(current), "recent_colors": _最近の色(),
        "form_title": "📝 イベントを編集する" if c else "✨ 新しいイベントを追加する",
    })


@owner_required
def calendar_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        if not d["dates"] and not d["date"]:
            messages.error(request, "日付を少なくとも1つ選択（追加）してください。")
            return _画面(request, None, request.POST, [])
        答 = admin_calendar.足す(d)
        if 答.get("status") == "ok":
            messages.success(request, "新しいイベントを追加しました！ 🎉")
            _通知(request, 答, d)
            return redirect("manage:calendar_list")
        messages.error(request, 答.get("message") or "追加に失敗しました。")
        return _画面(request, None, request.POST, [])
    return _画面(request, None, {
        "color": DEFAULT_COLOR, "notice_refresh": "on", "date": timezone.localdate().isoformat(),
    }, [])


@owner_required
def calendar_edit(request, row: int):
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, c)
        答 = admin_calendar.書き換える(d)
        if 答.get("status") == "ok":
            # 「お知らせ一覧に反映する」にチェックがあれば、掲載日時を今にして掲載順を更新する。
            if d.get("noticeStatus") == "公開":
                admin_calendar.一覧掲載を変える({"rowIdx": row, "status": "公開"})
            messages.success(request, "イベント情報を更新しました！")
            _通知(request, 答, d)
            return redirect(_戻る(c))
        messages.error(request, 答.get("message") or "追加に失敗しました。")
    values = request.POST or {
        "date": c.event_on.isoformat() if c.event_on else "", "title": c.title, "category": c.category,
        "menuRowIdx": c.menu_row or "", "desc": c.detail, "color": c.color or DEFAULT_COLOR,
        "publishAt": timezone.localtime(c.publish_at).strftime("%Y-%m-%dT%H:%M") if c.publish_at else "",
        "linkUrl": c.link_url, "linkButtonText": c.button_text,
        # 編集時は既定でオフ（軽微な修正で掲載順が動かないように）。旧アプリと同じ。
        "notice_refresh": "",
    }
    return _画面(request, c, values, _画像(c.image_url))


def _戻る(c):
    c.refresh_from_db()
    if c.event_on:
        return f"/manage/calendar/?year={c.event_on.year}&month={c.event_on.month}"
    return "/manage/calendar/"


@owner_required
@require_POST
def calendar_status(request, row: int):
    """一覧の「公開設定」（公開 / 非公開）。旧アプリの updateCalendarStatus。"""
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    status = "非公開" if request.POST.get("status") == "非公開" else "公開"
    答 = admin_calendar.公開を変える({"rowIdx": row, "status": status})
    if 答.get("status") == "ok":
        messages.success(request, "公開設定を変更しました")
    else:
        messages.error(request, 答.get("message") or "公開設定を変更できませんでした。")
    return redirect(request.POST.get("next") or "manage:calendar_list")


@owner_required
@require_POST
def calendar_delete(request, row: int):
    c = get_object_or_404(CalendarEvent, sheet_row=row, deleted=False)
    答 = admin_calendar.消す({"rowIdx": row, "reason": f"管理画面から（{request.user.username}）"})
    if 答.get("status") == "ok":
        messages.success(request, "イベントを削除しました")
    else:
        messages.error(request, "削除に失敗しました")
    return redirect(request.POST.get("next") or "manage:calendar_list")


@owner_required
@require_POST
def calendar_bulk_delete(request):
    rows = [int(x) for x in request.POST.getlist("rows") if x.isdigit()]
    if not rows:
        messages.error(request, "削除する項目を選択してください。")
        return redirect(request.POST.get("next") or "manage:calendar_list")
    答 = admin_calendar.まとめて消す({"rowIdxs": rows, "reason": f"管理画面から一括（{request.user.username}）"})
    messages.success(request, f"{答.get('deleted', 0)}件のイベントを削除しました")
    return redirect(request.POST.get("next") or "manage:calendar_list")
