"""メニュー管理（ホームに出る教室・施術の一覧）。

読み書きは gasapi/admin_menu.py（GAS の転送先と同じ）。**既定は非公開**（GAS と同じ）。
並べ替えは「行の中身の入れ替え」（admin_menu.動かす）。行番号は動かない。
"""

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Category, Menu
from apps.gasapi import admin_menu

from . import images, push

BOOKING_STATUSES = ["予約受付中", "予約対象外"]


def _候補の分類():
    名 = set(Category.objects.filter(kind="メニュー").values_list("name", flat=True))
    名 |= set(Menu.objects.filter(deleted=False).exclude(category="").values_list("category", flat=True))
    return sorted(n for n in 名 if n)


@login_required
def menu_list(request):
    rows = Menu.objects.filter(deleted=False).order_by("sheet_row")
    published = [m for m in rows if m.published]
    drafts = [m for m in rows if not m.published]
    return render(request, "manage/menu_list.html", {"published": published, "drafts": drafts})


def _入力(request, m: Menu | None):
    """フォームの値を、admin_menu が受ける形（GAS の payload と同じ名前）にする。"""
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "menus") for f in request.FILES.getlist("images")]
    status = request.POST.get("status") or "非公開"
    return {
        "rowIdx": m.sheet_row if m else 0,
        "name": request.POST.get("name", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "description": request.POST.get("description", ""),
        "reservationStatus": request.POST.get("reservationStatus", "予約対象外"),
        "publishStatus": status,
        "noticeStatus": "公開" if request.POST.get("notice_listed") else "非公開",
        "publishAt": request.POST.get("publishAt", ""),
        "imageUrls": kept + added,
        "imageUrl": (kept + added)[0] if (kept + added) else "",
    }


def _通知(request, 答, d):
    """GAS の _メニューの通知を送る_ と同じ条件。"""
    if not request.POST.get("send_push"):
        return
    状態 = 答.get("effectiveStatus") or d.get("publishStatus") or "非公開"
    公開日時 = 答.get("effectivePublishAt") or ""
    if 状態 != "公開":
        return
    if 公開日時:
        from apps.gasapi.admin_news import _日時

        t = _日時(公開日時)
        if t and t > timezone.now():
            return
    if push.全員へ送る("🍴 " + (d.get("name") or "ホーム更新"), "ホームのメニュー一覧が更新されました", page="home"):
        messages.info(request, "お客様のアプリへ通知を送りました。")
    else:
        messages.error(request, "通知を送れませんでした（メニューは保存されています）。")


@login_required
def menu_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        答 = admin_menu.足す(d)
        if 答.get("status") == "ok":
            messages.success(request, "メニューを追加しました。" if d["publishStatus"] == "公開" else "下書きとして保存しました。")
            _通知(request, 答, d)
            return redirect("manage:menu_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    return render(request, "manage/menu_form.html", {
        "menu": None, "existing_images": [], "categories": _候補の分類(),
        "booking_statuses": BOOKING_STATUSES, "values": request.POST or {"reservationStatus": "予約対象外"},
    })


@login_required
def menu_edit(request, row: int):
    m = get_object_or_404(Menu, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, m)
        答 = admin_menu.書き換える(d)
        if 答.get("status") == "ok":
            messages.success(request, "メニューを更新しました。")
            _通知(request, 答, d)
            return redirect("manage:menu_list")
        messages.error(request, 答.get("message") or "保存できませんでした。")
    values = request.POST or {
        "name": m.name, "category": m.category, "description": m.summary,
        "reservationStatus": m.booking_status or "予約対象外",
        "publishAt": timezone.localtime(m.publish_at).strftime("%Y-%m-%dT%H:%M") if m.publish_at else "",
        "notice_listed": "on" if m.notice_listed else "",
    }
    return render(request, "manage/menu_form.html", {
        "menu": m, "existing_images": list(m.image_urls or []), "categories": _候補の分類(),
        "booking_statuses": BOOKING_STATUSES, "values": values,
    })


@login_required
@require_POST
def menu_toggle(request, row: int):
    m = get_object_or_404(Menu, sheet_row=row, deleted=False)
    答 = admin_menu.書き換える({"rowIdx": row, "publishStatus": "非公開" if m.published else "公開"})
    if 答.get("status") == "ok":
        messages.success(request, f"「{m.name}」を{'非公開' if m.published else '公開'}にしました。")
    return redirect("manage:menu_list")


@login_required
@require_POST
def menu_move(request, row: int):
    答 = admin_menu.動かす({"rowIdx": row, "direction": request.POST.get("direction", "up")})
    if 答.get("status") != "ok":
        messages.error(request, "これ以上は動かせません。")
    return redirect("manage:menu_list")


@login_required
@require_POST
def menu_delete(request, row: int):
    m = get_object_or_404(Menu, sheet_row=row, deleted=False)
    答 = admin_menu.消す({"rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, f"「{m.name}」を削除しました。")
    return redirect("manage:menu_list")
