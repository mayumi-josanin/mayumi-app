"""Discord Webhook 通知（KEM の apps/devkanri/notifications.py の写し）。

プロジェクトに Webhook URL が入っていれば、タスクの割当とステータス変更を Discord に流す。
**送れなくても保存は成立させる。**通知のために画面が止まると、タスクを書けなくなるため。
KEM はスレッドで送っていたが、こちらは同期で送る（timeout 5秒。試験では urlopen を差し替える）。
KEM の「作業員の Discord ID でメンション」は、こちらに作業員が無いので外した。
"""

import json
import logging
from urllib.request import Request, urlopen

logger = logging.getLogger(__name__)


def _send_webhook(webhook_url, payload):
    try:
        data = json.dumps(payload).encode("utf-8")
        req = Request(webhook_url, data=data, headers={"Content-Type": "application/json"}, method="POST")
        urlopen(req, timeout=5)
    except Exception:
        # Webhook URL は知っていれば誰でも書き込める合鍵に近い。ログには先頭だけ残す
        logger.exception("Discord webhook送信に失敗: %s...", webhook_url[:40])


def _名前(user):
    return (user.get_full_name() or user.username) if user else "未割当"


def notify_task_assigned(task, assigned_by):
    """タスクが担当者に割り当てられた際に Discord 通知を送信。"""
    webhook_url = task.project.discord_webhook_url
    if not webhook_url or not task.assignee:
        return

    priority_emoji = {"low": "", "medium": "", "high": "!", "critical": "!!"}
    p_mark = priority_emoji.get(task.priority, "")
    due = str(task.due_date) if task.due_date else "未設定"

    embed = {
        "title": f"#{task.pk} {task.title}",
        "color": 0x1A2744,
        "fields": [
            {"name": "担当者", "value": _名前(task.assignee), "inline": True},
            {"name": "優先度", "value": f"{task.get_priority_display()} {p_mark}", "inline": True},
            {"name": "カテゴリ", "value": task.get_category_display(), "inline": True},
            {"name": "ステータス", "value": task.get_status_display(), "inline": True},
            {"name": "期限", "value": due, "inline": True},
            {"name": "プロジェクト", "value": task.project.name, "inline": True},
        ],
        "footer": {"text": f"割当者: {_名前(assigned_by)}"},
    }
    if task.description:
        embed["description"] = task.description[:200]

    _send_webhook(webhook_url, {"content": "**タスクが割り当てられました**", "embeds": [embed]})


def notify_task_status_changed(task, changed_by, old_status):
    """タスクのステータスが変更された際に Discord 通知を送信。"""
    webhook_url = task.project.discord_webhook_url
    if not webhook_url:
        return

    old_label = dict(task.Status.choices).get(old_status, old_status)
    embed = {
        "title": f"#{task.pk} {task.title}",
        "color": 0x2C4A7C,
        "fields": [
            {"name": "ステータス変更", "value": f"{old_label} → {task.get_status_display()}", "inline": False},
            {"name": "担当者", "value": _名前(task.assignee), "inline": True},
            {"name": "変更者", "value": _名前(changed_by), "inline": True},
        ],
    }
    _send_webhook(webhook_url, {"content": "**ステータスが変更されました**", "embeds": [embed]})
