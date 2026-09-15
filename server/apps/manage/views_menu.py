"""メニュー管理（ホームに出る教室・施術の一覧）。**中身は旧管理アプリ（#page-menus メニュー管理）と同じ。**

読み書きは gasapi/admin_menu.py（GAS の転送先と同じ）。**既定は非公開**（GAS と同じ）。

旧管理アプリの一覧は「下書き保存一覧」「公開済みメニュー一覧」の2つ。集計カード・検索・
絞り込み・一括削除・一覧からの公開切替は**無い**ので、ここにも置かない。
並べ替えは行をつまんで動かす（ドラッグ）。落としたところで表示順を保存する。
"""

import re

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.html import escape
from django.utils.safestring import mark_safe
from django.views.decorators.http import require_POST

from apps.content.models import Category, Menu
from apps.gasapi import admin_menu

from . import images, push
from .permissions import owner_required

# 旧管理アプリの select と同じ並び。先頭（予約受付中）が新規追加のときの既定。
BOOKING_STATUSES = ["予約受付中", "予約対象外"]


def _候補の分類():
    名 = set(Category.objects.filter(kind="メニュー").values_list("name", flat=True))
    名 |= set(Menu.objects.filter(deleted=False).exclude(category="").values_list("category", flat=True))
    return sorted(n for n in 名 if n)


def _整形(text: str):
    """旧管理アプリの renderAdminFormattedTextHtml と同じ。

    太字（<strong>/<b>）と下線（<u>）だけ残し、他はそのまま文字として出す。
    「📷 https://…」の行（説明に貼った画像の印）は一覧では消す。改行は <br>。
    """
    s = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"(^|\n)\s*📷\s*https?://\S+", r"\1", s)
    s = re.sub(r"<\s*(strong|b)\s*>", "[[B_OPEN]]", s, flags=re.I)
    s = re.sub(r"<\s*/\s*(strong|b)\s*>", "[[B_CLOSE]]", s, flags=re.I)
    s = re.sub(r"<\s*u\s*>", "[[U_OPEN]]", s, flags=re.I)
    s = re.sub(r"<\s*/\s*u\s*>", "[[U_CLOSE]]", s, flags=re.I)
    s = escape(s)
    s = (s.replace("[[B_OPEN]]", "<strong>").replace("[[B_CLOSE]]", "</strong>")
          .replace("[[U_OPEN]]", "<u>").replace("[[U_CLOSE]]", "</u>"))
    return mark_safe(s.replace("\n", "<br>"))


def _更新日(m: Menu) -> str:
    """旧管理アプリの formatDisplayDate(updatedAt || date)。日付だけ（YYYY/MM/DD）、無ければ「ー」。"""
    if m.updated_at:
        return timezone.localtime(m.updated_at).strftime("%Y/%m/%d")
    if m.registered_on:
        return m.registered_on.strftime("%Y/%m/%d")
    return "ー"


def _一件(m: Menu) -> dict:
    画像 = list(m.image_urls or [])
    return {
        "m": m, "image": 画像[0] if 画像 else "",
        "category": m.category or "未設定",
        "description": _整形(m.summary or "概要説明は未入力です"),
        "updated": _更新日(m),
    }


def _お客様と同じ順(rows):
    """お客様アプリ（app.js）がホームで並べるのと同じ: 表示順の大きい順、同じなら更新日時の新しい順。

    管理画面でもこの順に出す。つまんで動かした結果がそのまま「お客様に見える順」になる。
    """
    def 鍵(m):
        更新 = m.updated_at.timestamp() if m.updated_at else 0
        return (-(m.sort_key or 0), -更新, m.sheet_row)
    return sorted(rows, key=鍵)


@owner_required
def menu_list(request):
    rows = _お客様と同じ順(Menu.objects.filter(deleted=False))
    return render(request, "manage/menu_list.html", {
        "drafts": [_一件(m) for m in rows if not m.published],
        "published": [_一件(m) for m in rows if m.published],
    })


def _入力(request, m: Menu | None):
    """フォームの値を、admin_menu が受ける形（GAS の payload と同じ名前）にする。"""
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "menus") for f in request.FILES.getlist("images")]
    status = request.POST.get("status") or "公開"
    d = {
        "rowIdx": m.sheet_row if m else 0,
        "name": request.POST.get("name", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "description": request.POST.get("description", ""),
        "reservationStatus": request.POST.get("reservationStatus", "予約受付中"),
        "publishStatus": status,
        "publishAt": request.POST.get("publishAt", ""),
        "imageUrls": kept + added,
        "imageUrl": (kept + added)[0] if (kept + added) else "",
    }
    # 「この保存をお知らせ一覧に反映する」。新規はチェックのとおりに載せる／載せない。
    # 編集のときは、ここでは触らない（チェックしたときだけ、保存後に載せ直す）。
    if m is None:
        d["noticeStatus"] = "公開" if request.POST.get("notice_refresh") else "非公開"
    return d


def _通知(request, 答, d):
    """GAS の _メニューの通知を送る_ と同じ条件。公開保存のときだけ送る。"""
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


def _画面(request, m, values, existing):
    return render(request, "manage/menu_form.html", {
        "menu": m, "values": values, "existing_images": existing,
        "categories": _候補の分類(), "booking_statuses": BOOKING_STATUSES,
        "form_title": "📝 メニューを編集する" if m else "✨ 新しいメニューを追加する",
    })


def _足りない(request, d) -> bool:
    """旧管理アプリの submitMenu と同じ確認。"""
    if d["category"] and d["name"]:
        return False
    messages.error(request, "カテゴリとメニュー名を入力してください")
    return True


@owner_required
def menu_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        if not _足りない(request, d):
            答 = admin_menu.足す(d)
            if 答.get("status") == "ok":
                messages.success(request, "メニューを追加しました")
                _通知(request, 答, d)
                return redirect("manage:menu_list")
            messages.error(request, 答.get("message") or "保存に失敗しました")
    # 新規追加は「お知らせ一覧に反映する」が既定でオン（旧管理アプリと同じ）
    values = request.POST or {"reservationStatus": "予約受付中", "notice_refresh": "on"}
    return _画面(request, None, values, [])


@owner_required
def menu_edit(request, row: int):
    m = get_object_or_404(Menu, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, m)
        if not _足りない(request, d):
            答 = admin_menu.書き換える(d)
            if 答.get("status") == "ok":
                if request.POST.get("notice_refresh"):
                    # チェックしたときだけ、お知らせ一覧に載せ直す（掲載日時が今になり、上に来る）
                    admin_menu.一覧掲載を変える({"rowIdx": row, "status": "公開"})
                messages.success(request, "メニューを更新しました")
                _通知(request, 答, d)
                return redirect("manage:menu_list")
            messages.error(request, 答.get("message") or "保存に失敗しました")
    # 編集のときは「お知らせ一覧に反映する」はオフから（旧管理アプリの editMenu と同じ）
    values = request.POST or {
        "name": m.name, "category": m.category, "description": m.summary,
        "reservationStatus": m.booking_status or "予約受付中",
        "publishAt": timezone.localtime(m.publish_at).strftime("%Y-%m-%dT%H:%M") if m.publish_at else "",
    }
    return _画面(request, m, values, list(m.image_urls or []))


@owner_required
@require_POST
def menu_order(request):
    """つまんで動かした順で表示順を保存する（旧管理アプリの saveMenuOrder と同じ計算）。

    rows は上から順の行番号。上ほど大きい数を入れる（お客様のホームは大きい順）。
    """
    rows = [int(x) for x in request.POST.getlist("rows") if x.isdigit()]
    count = len(rows)
    答 = admin_menu.並び順を書く({"updates": [
        {"rowIdx": r, "sortOrder": (count - i) * 1000 + 1000000} for i, r in enumerate(rows)
    ]})
    if 答.get("status") == "ok":
        messages.success(request, "メニュー順序を保存しました")
    else:
        messages.error(request, "保存に失敗しました: " + (答.get("message") or ""))
    return redirect("manage:menu_list")


@owner_required
@require_POST
def menu_delete(request, row: int):
    m = get_object_or_404(Menu, sheet_row=row, deleted=False)
    答 = admin_menu.消す({"rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, f"「{m.name}」を削除しました")
    else:
        messages.error(request, "削除に失敗しました")
    return redirect("manage:menu_list")
