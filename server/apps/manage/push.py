"""管理画面からのプッシュ通知。中身は gasapi/admin_push.py（GAS の転送先と同じ）。

旧管理アプリ経由でも Django 経由でも、記録は同じ表（PushNotice）に同じ形で残る。
"""

from apps.content import onesignal
from apps.gasapi import admin_push


def 設定済みか() -> bool:
    return onesignal.設定済みか()


def 全員へ送る(title: str, body: str, *, page: str = "home") -> bool:
    """投稿に伴う自動送信（GAS の sendAutoPush と同じ）。送れなければ False。"""
    if not 設定済みか():
        return False
    答 = admin_push.送る({"mode": "auto", "title": title, "body": body,
                          "targetStatus": "all", "targetPage": page, "previewBody": body})
    return bool(答.get("sent"))
