"""プッシュ通知の表（PUSH_NOTICES）を、GAS の代わりに読み書きし、送る。

**GAS が転送してくる先**（SERVER_TABLES に push を入れたとき）。
GAS の broadcastPush / sendAutoPush / processScheduledPushQueue / handleDeleteRow(PUSH) と
同じ振る舞いにしてある。Django の管理画面（apps/manage）もここを呼ぶ。

## 行番号を鍵として使い続ける

旧管理アプリの通知の履歴は rowIdx で行を指す（削除など）。**そこを変えない。**

## 送る前に必ず記録を書く

GAS と同じ順（先に行を書き、送れたら sentAt・通知ID・結果を書き足す）。
先に送って後で書くと、送れたのに記録が無い状態が起きる。
送れなかったときは「送信失敗」と理由を残す（GAS の sendAutoPush は失敗を握りつぶして
「自動送信済み」のまま残していた。それで4月末からの失敗に誰も気づけなかった）。
"""

import json

from django.db import transaction
from django.utils import timezone

from apps.content import onesignal
from apps.content.models import PushNotice

from .admin_news import _日時, _文

SENT = "送信済み"
AUTO_SENT = "自動送信済み"
DRAFT = "下書き"
SCHEDULED = "予約済み"
FAILED = "送信失敗"


def _対象者(d) -> list:
    """GAS の parsePushTargetUsers_。targetUsers か、targetDetail（JSON の配列）から届け先を取る。"""
    users = d.get("targetUsers")
    if not isinstance(users, list) or not users:
        生 = _文(d.get("targetDetail")).strip()
        users = []
        if 生.startswith("["):
            try:
                users = json.loads(生)
            except ValueError:
                users = []
    出 = []
    for u in users if isinstance(users, list) else []:
        if isinstance(u, dict):
            出.append({
                "memberId": _文(u.get("memberId")).strip(),
                "name": _文(u.get("name")).strip(),
                "subscription": _文(u.get("subscription") or u.get("subscriptionId")).strip(),
            })
    return 出


def _届け先の端末(d, 記録: PushNotice) -> list:
    """GAS の buildOneSignalPayload_ と同じ条件で、端末を絞るか全員かを決める。"""
    ids = [u["subscription"] for u in _対象者(d) if u.get("subscription")]
    detail = (記録.target_detail or "").strip()
    絞る = bool(ids) and ((記録.target_status or "all") != "all" or detail.startswith("テスト送信:"))
    return ids if 絞る else []


def _記録を書く(d, status: str) -> PushNotice:
    """GAS の buildPushNoticeRowFromInput_ + writePushNoticeRow_。rowIdx があれば上書き、無ければ追加。"""
    users = _対象者(d)
    detail = _文(d.get("targetDetail")).strip() or (json.dumps(users, ensure_ascii=False) if users else "")
    値 = {
        "title": _文(d.get("title")).strip(),
        "body": _文(d.get("body")).strip(),
        "target_status": _文(d.get("targetStatus")).strip() or "all",
        "target_detail": detail,
        "recipient_count": int(d.get("recipientCount") or len(users) or 0) or None,
        "status": status,
        "scheduled_at": _日時(d.get("scheduledAt")),
        "target_page": onesignal.ページ(d.get("targetPage")),
        "preview_body": _文(d.get("previewBody")).strip() or _文(d.get("body")).strip(),
        "updated_at": timezone.now(),
    }
    行 = int(d.get("rowIdx") or 0)
    with transaction.atomic():
        既存 = PushNotice.objects.select_for_update().filter(sheet_row=行).first() if 行 > 1 else None
        if 既存:
            for k, v in 値.items():
                setattr(既存, k, v)
            既存.save()
            return 既存
        最大 = (
            PushNotice.objects.select_for_update().order_by("-sheet_row")
            .values_list("sheet_row", flat=True).first() or 1
        )
        return PushNotice.objects.create(sheet_row=最大 + 1, **値)


def _送って記録する(記録: PushNotice, ids: list, 結果の字: str, 成功の状態: str) -> tuple[bool, str]:
    try:
        答 = onesignal.送信(onesignal.内容(記録.title, 記録.body, 記録.target_page, ids))
    except onesignal.送信できない as e:
        記録.sent_at = timezone.now()
        記録.status = FAILED
        記録.result = str(e)[:1000]
        記録.save(update_fields=["sent_at", "status", "result", "changed_at"])
        return False, str(e)
    記録.sent_at = timezone.now()
    記録.status = 成功の状態
    記録.notification_id = _文(答.get("id"))[:64]
    受取 = int(答.get("recipients") or 0)
    if 受取:
        記録.recipient_count = 受取
    記録.result = 結果の字
    記録.save(update_fields=["sent_at", "status", "notification_id", "recipient_count", "result", "changed_at"])
    return True, 結果の字


def 送る(d):
    """GAS の broadcastPush（mode: send / draft / schedule / test）＋ sendAutoPush（mode: auto）。"""
    d = d or {}
    mode = _文(d.get("mode")).strip() or "send"
    if not _文(d.get("title")).strip() and mode != "draft":
        return {"status": "error", "message": "タイトルを入力してください"}

    if mode == "draft":
        記録 = _記録を書く(d, DRAFT)
        return {"status": "ok", "message": "通知を下書き保存しました", "rowIdx": 記録.sheet_row,
                "recipientCount": 記録.recipient_count or 0}

    if mode == "schedule":
        記録 = _記録を書く(d, SCHEDULED)
        if not 記録.scheduled_at:
            return {"status": "error", "message": "配信予定日時を入力してください"}
        return {"status": "ok", "message": "通知を予約しました", "rowIdx": 記録.sheet_row,
                "recipientCount": 記録.recipient_count or 0}

    if mode == "test":
        users = _対象者(d)
        if not users:
            return {"status": "error", "message": "テスト送信先を選択してください"}
        記録 = _記録を書く(d, SENT)
        ok, 結果 = _送って記録する(記録, [u["subscription"] for u in users if u.get("subscription")],
                              "テスト送信済み", SENT)
        if not ok:
            return {"status": "error", "message": 結果, "rowIdx": 記録.sheet_row}
        return {"status": "ok", "message": "テスト通知を送信しました", "rowIdx": 記録.sheet_row,
                "recipientCount": len(users)}

    if mode == "auto":
        # お知らせ・カレンダーなどの投稿に伴う自動送信。**失敗しても投稿は成立させる**（error を返さない）。
        記録 = _記録を書く(d, AUTO_SENT)
        ok, 結果 = _送って記録する(記録, [], AUTO_SENT, AUTO_SENT)
        return {"status": "ok" if ok else "error", "message": 結果, "rowIdx": 記録.sheet_row,
                "recipientCount": 記録.recipient_count or 0, "sent": ok}

    記録 = _記録を書く(d, SENT)
    ok, 結果 = _送って記録する(記録, _届け先の端末(d, 記録), "送信済み", SENT)
    if not ok:
        return {"status": "error", "message": 結果, "rowIdx": 記録.sheet_row}
    return {"status": "ok", "message": "通知を送信しました", "rowIdx": 記録.sheet_row,
            "recipientCount": 記録.recipient_count or 0}


def 予約分を送る(d=None):
    """GAS の processScheduledPushQueue。予約済みで時刻が来たものを送る。**5分ごとに GAS の
    トリガーが呼んでくる**（時計は GAS のまま。サーバーに cron を足さない）。"""
    いま = timezone.now()
    処理 = 0
    for 記録 in PushNotice.objects.filter(status=SCHEDULED, deleted=False).order_by("sheet_row"):
        # GAS の isPublishAtAvailable_: 予定が空なら「来ている」扱い
        if 記録.scheduled_at and 記録.scheduled_at > いま:
            continue
        ids = _届け先の端末({"targetDetail": 記録.target_detail}, 記録)
        _送って記録する(記録, ids, "予約送信済み", SENT)
        処理 += 1
    return {"status": "ok", "processed": 処理}


def 消す(d):
    """GAS の handleDeleteRow（sheet: PUSH）。印を付けるだけ。"""
    行 = int((d or {}).get("rowIdx") or 0)
    p = PushNotice.objects.filter(sheet_row=行, deleted=False).first() if 行 else None
    if not p:
        return {"status": "error", "message": "通知が見つかりませんでした"}
    p.deleted = True
    p.deleted_at = timezone.now()
    p.save(update_fields=["deleted", "deleted_at", "changed_at"])
    return {"status": "ok"}


def まとめて消す(d):
    番号 = (d or {}).get("rowIdxs") or (d or {}).get("rows") or []
    if not isinstance(番号, list):
        return {"status": "error", "message": "行番号の並びが必要です"}
    数 = sum(1 for r in 番号 if 消す({"rowIdx": r}).get("status") == "ok")
    return {"status": "ok", "deleted": 数}
