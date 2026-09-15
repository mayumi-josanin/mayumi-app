"""通知管理。お客様アプリへのプッシュ通知の履歴と、手で送る画面。

旧管理アプリの通知管理は履歴と削除だけになっていた（送る画面が無くなっていた）。
ここでは「今すぐ全員に送る／下書き／予約」を戻す。
読み書きは gasapi/admin_push.py（GAS の転送先と同じ）。
"""

from django.contrib import messages
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from apps.content import onesignal
from apps.content.models import PushNotice
from apps.gasapi import admin_push

from .permissions import owner_required

PAGE_LABELS = [
    ("home", "ホーム"), ("news", "NEWS"), ("calendar", "カレンダー"), ("shop", "ショップ"),
    ("notices", "通知一覧"), ("mypage", "マイページ"), ("menu-list", "メニュー"),
]
STATUS_BADGE = {
    admin_push.SENT: "badge-green", admin_push.AUTO_SENT: "badge-green",
    admin_push.SCHEDULED: "badge-blue", admin_push.DRAFT: "badge-gray", admin_push.FAILED: "badge-red",
}


@owner_required
def push_list(request):
    rows = PushNotice.objects.filter(deleted=False).order_by("-sheet_row")[:200]
    items = [{"p": p, "badge": STATUS_BADGE.get(p.status or "", "badge-gray")} for p in rows]
    return render(request, "manage/push_list.html", {
        "items": items,
        "pages": PAGE_LABELS,
        "configured": onesignal.設定済みか(),
        "draft": {"title": request.GET.get("title", ""), "body": request.GET.get("body", "")},
    })


@owner_required
@require_POST
def push_send(request):
    mode = request.POST.get("mode") or "send"
    if mode not in ("send", "draft", "schedule"):
        mode = "send"
    答 = admin_push.送る({
        "mode": mode,
        "rowIdx": request.POST.get("rowIdx") or 0,
        "title": request.POST.get("title", ""),
        "body": request.POST.get("body", ""),
        "targetStatus": "all",
        "targetPage": request.POST.get("targetPage", "home"),
        "scheduledAt": request.POST.get("scheduledAt", ""),
    })
    if 答.get("status") == "ok":
        messages.success(request, 答.get("message") or "送りました。")
    else:
        messages.error(request, 答.get("message") or "送れませんでした。")
    return redirect("manage:push_list")


@owner_required
@require_POST
def push_delete(request, row: int):
    答 = admin_push.消す({"rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, "通知の記録を削除しました。")
    else:
        messages.error(request, 答.get("message") or "削除できませんでした。")
    return redirect("manage:push_list")
