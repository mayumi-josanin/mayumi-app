"""お客様アプリへのプッシュ通知（OneSignal）を、このサーバーから送る。

GAS の sendAutoPush / buildOneSignalPayload_ / sendPushNoticePayload_ と同じ中身。
旧管理アプリから投稿すると GAS が送るが、この管理画面からの投稿は GAS を
通らないので、ここで送る。

送るのに要るもの（.env）: ONESIGNAL_APP_ID / ONESIGNAL_REST_API_KEY
（GAS のスクリプトプロパティと同じ値）。**無ければ送らず False を返す。**

送った記録は PushNotice の表に「自動送信済み」で残す（GAS は通知シートに残す）。
"""

import json
import logging
import urllib.error
import urllib.request

from django.conf import settings
from django.db import transaction
from django.utils import timezone

from apps.content.models import PushNotice

logger = logging.getLogger(__name__)

ONESIGNAL_URL = "https://onesignal.com/api/v1/notifications"
# お客様アプリが登録時に付けるタグ。GAS の ONE_SIGNAL_APP_SCOPE_KEY/VALUE と同じ。
APP_SCOPE_KEY = "app_scope"
APP_SCOPE_VALUE = "mayumi_josanin_app"
PAGES = ["home", "shop", "calendar", "news", "notices", "mypage", "menu-list"]
STATUS_AUTO_SENT = "自動送信済み"
STATUS_FAILED = "送信失敗"


def 設定済みか() -> bool:
    return bool(settings.ONESIGNAL_APP_ID and settings.ONESIGNAL_REST_API_KEY)


def _page(value: str) -> str:
    v = (value or "").strip().lower()
    return v if v in PAGES else "home"


def _payload(title: str, body: str, page: str) -> dict:
    return {
        "app_id": settings.ONESIGNAL_APP_ID,
        "contents": {"en": body, "ja": body},
        "headings": {"en": title, "ja": title},
        "url": "https://mayumi-josanin.github.io/mayumi-app/?open=" + page,
        "data": {"openPage": page},
        "filters": [{"field": "tag", "key": APP_SCOPE_KEY, "relation": "=", "value": APP_SCOPE_VALUE}],
    }


def _送信(payload: dict) -> dict:
    req = urllib.request.Request(
        ONESIGNAL_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": "Basic " + settings.ONESIGNAL_REST_API_KEY,
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=20) as res:
        return json.loads(res.read().decode("utf-8") or "{}")


def 全員へ送る(title: str, body: str, *, page: str = "home") -> bool:
    """全員（アプリのタグを持つ端末）へ1通送り、記録を残す。送れなければ False。**例外にしない。**"""
    if not 設定済みか():
        logger.warning("OneSignal の設定が無いので、通知は送りません")
        return False
    page = _page(page)
    with transaction.atomic():
        最大 = (
            PushNotice.objects.select_for_update().order_by("-sheet_row")
            .values_list("sheet_row", flat=True).first() or 1
        )
        記録 = PushNotice.objects.create(
            sheet_row=最大 + 1,
            title=title.strip(),
            body=body.strip(),
            target_status="all",
            target_page=page,
            preview_body=body.strip(),
            status=STATUS_AUTO_SENT,
            updated_at=timezone.now(),
        )
    try:
        答 = _送信(_payload(title, body, page))
        記録.sent_at = timezone.now()
        記録.notification_id = str(答.get("id") or "")
        記録.recipient_count = int(答.get("recipients") or 0) or None
        記録.result = STATUS_AUTO_SENT
        記録.save(update_fields=["sent_at", "notification_id", "recipient_count", "result", "changed_at"])
        return True
    except (urllib.error.URLError, urllib.error.HTTPError, ValueError, OSError) as e:
        logger.exception("OneSignal へ送れませんでした")
        記録.status = STATUS_FAILED
        記録.result = f"{STATUS_FAILED}: {e}"[:1000]
        記録.save(update_fields=["status", "result", "changed_at"])
        return False
