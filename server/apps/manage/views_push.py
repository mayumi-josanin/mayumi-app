"""通知管理。**中身は旧管理アプリ（#page-users 後半「📣 全員へPush通知を送信」と
「📣 送信済みお知らせ（Push通知）履歴」）と同じ。**

送る画面の項目は旧管理アプリの送信フォーム（通知タイトル・通知本文・配信対象・開く画面・
配信予定日時・テスト送信先、プレビュー／下書き保存／テスト送信／予約保存／一斉送信）。
※ 旧管理アプリでは 2026-04-04 の商品特別価格の再適用（94c878c）でフォームの HTML だけが落ち、
  JS（savePushDraft / schedulePushNotice / sendPushTest / previewPushNotice）と
  プレビューの窓・テンプレートの「使う」はそのまま残っていた。**落ちる前の項目をそのまま出す。**

履歴は集計カード（通知総数・7日以内・現在表示中）、タイトル・本文の検索、
新しい順の一覧（送信日時・タイトル・本文・操作）、一括削除。
読み書きは gasapi/admin_push.py（GAS の転送先と同じ）。届け先は gasapi/admin_member.通知の届け先()。
"""

from datetime import timedelta

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content import onesignal
from apps.content.models import PushNotice
from apps.gasapi import admin_member, admin_push

from .permissions import owner_required

# 「開く画面」の選択肢（旧管理アプリ #pushTargetPage と同じ並び・同じ文言）
PAGE_LABELS = [
    ("home", "ホーム"), ("shop", "ショップ"), ("calendar", "カレンダー"), ("news", "NEWS"),
    ("notices", "お知らせ一覧"), ("mypage", "マイページ"), ("menu-list", "メニュー一覧"),
]

# 「配信対象」の選択肢（旧管理アプリ #pushAudienceMode / buildSelectedPushAudience と同じ）
AUDIENCE_LABELS = [
    ("all", "Push許可中の全員"),
    ("pending-orders", "受付中注文がある会員"),
    ("reward-available", "未使用特典がある会員"),
    ("no-orders", "注文履歴がない会員"),
]


def _絞る(users: list, mode: str) -> list:
    """旧管理アプリの buildSelectedPushAudience と同じ条件で届け先を絞る。"""
    if mode == "pending-orders":
        return [u for u in users if int(u.get("pendingOrderCount") or 0) > 0]
    if mode == "reward-available":
        return [u for u in users if int(u.get("rewardCount") or 0) > 0]
    if mode == "no-orders":
        return [u for u in users if int(u.get("orderCount") or 0) == 0]
    return list(users)


def _届け先():
    """Push を許可している会員（getPushUsers と同じ中身）。"""
    答 = admin_member.通知の届け先()
    return list(答.get("users") or [])


def _日時の数(p: PushNotice) -> int:
    """並び順と「7日以内」の基準。GAS の getPushNotices と同じで sentAt → scheduledAt → updatedAt の順に使う。"""
    for d in (p.sent_at, p.scheduled_at, p.updated_at):
        if d:
            return int(d.timestamp() * 1000)
    return 0


def _表示日(d) -> str:
    """旧管理アプリの formatDisplayDate（日付だけ、空なら「ー」）。"""
    if not d:
        return "ー"
    return timezone.localtime(d).strftime("%Y/%m/%d")


def _検索の字(v: str) -> str:
    """旧管理アプリの normalizeFilterText（小文字・空白を除く・長音記号は「-」に）。"""
    出 = "".join(str(v or "").lower().split())
    for c in "‐－―ー":
        出 = 出.replace(c, "-")
    return 出


@owner_required
def push_list(request):
    q = (request.GET.get("q") or "").strip()
    全部 = [p for p in PushNotice.objects.filter(deleted=False, deleted_at__isnull=True)]
    全部.sort(key=lambda p: (-_日時の数(p), -p.sheet_row))
    見せる = []
    語 = _検索の字(q)
    for p in 全部:
        if 語:
            日 = next((d for d in (p.sent_at, p.scheduled_at, p.updated_at) if d), None)
            if 語 not in _検索の字(" ".join([_表示日(日), p.title or "", p.body or ""])):
                continue
        見せる.append({
            "p": p,
            "date": _表示日(p.sent_at or p.scheduled_at or p.updated_at),
            "scheduled": _表示日(p.scheduled_at) if p.scheduled_at else "",
            "status": p.status or "送信済み",
            "target_page": p.target_page or "home",
            "target_status": p.target_status or "all",
        })
    七日前 = (timezone.now() - timedelta(days=7)).timestamp() * 1000
    summary = [
        ("通知総数", len(全部)),
        ("7日以内", sum(1 for p in 全部 if _日時の数(p) >= 七日前)),
        ("現在表示中", len(見せる)),
    ]

    users = _届け先()
    件数 = {mode: len(_絞る(users, mode)) for mode, _ in AUDIENCE_LABELS}
    return render(request, "manage/push_list.html", {
        "items": 見せる, "summary": summary, "q": q,
        "meta": (f"{len(見せる)} / {len(全部)} 件を表示" if 全部 else "通知履歴はありません"),
        "empty": ("条件に一致する通知履歴はありません" if 全部 else "通知履歴がありません"),
        "pages": PAGE_LABELS,
        "audiences": [(mode, label, 件数[mode]) for mode, label in AUDIENCE_LABELS],
        "recipients": users,
        "configured": onesignal.設定済みか(),
        # 旧管理アプリはタイトルに「まゆみ助産院」が最初から入っている
        "draft": {"title": request.GET.get("title", "まゆみ助産院"), "body": request.GET.get("body", "")},
    })


def _会員だけ(u: dict) -> dict:
    return {"memberId": u.get("memberId"), "name": u.get("name"), "subscription": u.get("subscription")}


@owner_required
@require_POST
def push_send(request):
    """旧管理アプリの sendBroadcastPush / savePushDraft / schedulePushNotice / sendPushTest。"""
    mode = request.POST.get("mode") or "send"
    if mode not in ("send", "draft", "schedule", "test"):
        mode = "send"
    title = request.POST.get("title", "").strip()
    body = request.POST.get("body", "").strip()
    scheduled_at = request.POST.get("scheduledAt", "").strip()
    if not title or not body:
        messages.error(request, "タイトルと本文を入力してください")
        return redirect("manage:push_list")
    if mode == "schedule" and not scheduled_at:
        messages.error(request, "タイトル・本文・配信予定日時を入力してください")
        return redirect("manage:push_list")

    # 配信対象は旧管理アプリの collectPushNoticePayload と同じ形で渡す
    # （全員のときは会員の一覧を入れない。絞ったときだけ届け先を入れる）。
    audience_mode = request.POST.get("targetStatus") or "all"
    if audience_mode not in dict(AUDIENCE_LABELS):
        audience_mode = "all"
    users = _絞る(_届け先(), audience_mode)
    payload = {
        "mode": mode,
        "rowIdx": request.POST.get("rowIdx") or 0,
        "title": title,
        "body": body,
        "targetStatus": audience_mode,
        "targetDetail": dict(AUDIENCE_LABELS)[audience_mode],
        "targetUsers": [_会員だけ(u) for u in users] if audience_mode != "all" else [],
        "recipientCount": len(users),
        "targetPage": request.POST.get("targetPage", "home"),
        "scheduledAt": scheduled_at,
    }

    if mode == "test":
        # テスト送信先は「テスト送信先」で選んだ1人だけ（sendPushTest と同じ）
        member_id = request.POST.get("testRecipient", "").strip()
        target = next((u for u in _届け先() if str(u.get("memberId") or "") == member_id), None)
        if not target:
            messages.error(request, "テスト送信先を選択してください")
            return redirect("manage:push_list")
        payload.update({
            "targetUsers": [_会員だけ(target)],
            "targetDetail": f"テスト送信:{target.get('name') or target.get('memberId')}",
            "recipientCount": 1,
        })

    答 = admin_push.送る(payload)
    ok = 答.get("status") == "ok"
    if mode == "draft":
        messages.success(request, "下書きを保存しました") if ok else messages.error(request, "下書き保存に失敗しました")
    elif mode == "schedule":
        messages.success(request, "通知を予約しました") if ok else messages.error(request, "通知予約に失敗しました")
    elif mode == "test":
        messages.success(request, "テスト通知を送信しました") if ok else messages.error(request, "テスト送信に失敗しました")
    elif ok:
        messages.success(request, "送信完了しました！")
    else:
        messages.error(request, "送信エラー: " + (答.get("message") or "不明なエラー"))
    return redirect("manage:push_list")


@owner_required
@require_POST
def push_delete(request, row: int):
    答 = admin_push.消す({"rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, "通知を削除しました")
    else:
        messages.error(request, "削除に失敗しました: " + (答.get("message") or "ステータスが不適切です"))
    return redirect("manage:push_list")


@owner_required
@require_POST
def push_bulk_delete(request):
    rows = [int(x) for x in request.POST.getlist("rows") if x.isdigit()]
    if not rows:
        messages.error(request, "削除する項目を選択してください。")
        return redirect(request.POST.get("next") or "manage:push_list")
    答 = admin_push.まとめて消す({"rowIdxs": rows})
    if 答.get("status") == "ok":
        messages.success(request, f"{答.get('deleted', 0)}件の通知履歴を削除しました")
    else:
        messages.error(request, "削除に失敗しました: " + (答.get("message") or "ステータスが不適切です"))
    return redirect(request.POST.get("next") or "manage:push_list")
