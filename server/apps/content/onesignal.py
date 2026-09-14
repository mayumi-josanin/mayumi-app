"""OneSignal（お客様アプリへのプッシュ通知）の送り口。

GAS の buildOneSignalPayload_ / sendPushNoticePayload_ と同じ中身。
**鍵は .env（ONESIGNAL_APP_ID / ONESIGNAL_REST_API_KEY）。**GAS のスクリプトプロパティと同じ値。
記録（PushNotice の表）の読み書きはここではしない（gasapi/admin_push.py）。
"""

import json
import urllib.error
import urllib.request

from django.conf import settings

URL = "https://onesignal.com/api/v1/notifications"
# お客様アプリが登録時に付けるタグ。GAS の ONE_SIGNAL_APP_SCOPE_KEY/VALUE と同じ。
APP_SCOPE_KEY = "app_scope"
APP_SCOPE_VALUE = "mayumi_josanin_app"
PAGES = ["home", "shop", "calendar", "news", "notices", "mypage", "menu-list"]


class 送信できない(Exception):
    """OneSignal に届かなかった／断られた。"""


def 設定済みか() -> bool:
    return bool(settings.ONESIGNAL_APP_ID and settings.ONESIGNAL_REST_API_KEY)


def ページ(value) -> str:
    v = str(value or "").strip().lower()
    return v if v in PAGES else "home"


def 内容(title: str, body: str, page: str, subscription_ids=None) -> dict:
    """GAS の buildOneSignalPayload_。届け先が指定されていればその端末だけ、無ければ全員（タグ）。"""
    page = ページ(page)
    payload = {
        "app_id": settings.ONESIGNAL_APP_ID,
        "contents": {"en": body, "ja": body},
        "headings": {"en": title, "ja": title},
        "url": "https://mayumi-josanin.github.io/mayumi-app/?open=" + page,
        "data": {"openPage": page},
    }
    ids = [str(x).strip() for x in (subscription_ids or []) if str(x or "").strip()]
    if ids:
        payload["include_subscription_ids"] = ids
    else:
        payload["filters"] = [
            {"field": "tag", "key": APP_SCOPE_KEY, "relation": "=", "value": APP_SCOPE_VALUE}
        ]
    return payload


def 送信(payload: dict) -> dict:
    """送って OneSignal の答え（id / recipients）を返す。駄目なら 送信できない を投げる。"""
    if not 設定済みか():
        raise 送信できない("OneSignal の鍵が設定されていません")
    req = urllib.request.Request(
        URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": "Basic " + settings.ONESIGNAL_REST_API_KEY,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as res:
            答 = json.loads(res.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as e:
        本文 = ""
        try:
            本文 = e.read().decode("utf-8", "ignore")[:300]
        except Exception:
            pass
        raise 送信できない(f"OneSignal送信失敗: HTTP {e.code} {本文}") from e
    except (urllib.error.URLError, OSError, ValueError) as e:
        raise 送信できない(f"OneSignal送信失敗: {e}") from e
    if 答.get("errors"):
        raise 送信できない(f"OneSignal送信失敗: {答.get('errors')}")
    return 答
