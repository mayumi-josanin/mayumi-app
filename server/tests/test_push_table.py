"""プッシュ通知の表をサーバーで持つ。GAS の broadcastPush / 予約送信 / 削除と同じ振る舞い。

GAS は SERVER_TABLES に push を入れると、ここ（/api の broadcastPush など）へ転送してくる。
"""

import io
import json
from datetime import timedelta

import pytest
from django.utils import timezone

from apps.content import onesignal
from apps.content.models import PushNotice
from apps.gasapi import admin_push

pytestmark = pytest.mark.django_db

KEY = "api-key-for-test"


@pytest.fixture
def api(settings):
    settings.API_KEY = KEY
    settings.ONESIGNAL_APP_ID = "app-0001"
    settings.ONESIGNAL_REST_API_KEY = "key-0001"


@pytest.fixture
def fake_onesignal(monkeypatch):
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append(json.loads(req.data))
        return _Res(json.dumps({"id": "n-777", "recipients": 97}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


def _post(client, data):
    return client.post("/api", data=json.dumps(data), content_type="application/json",
                       HTTP_X_API_KEY=KEY).json()


# ========== GAS からの転送（/api） ==========


def test_合鍵が無ければ断る(client, api):
    r = client.post("/api", data=json.dumps({"type": "broadcastPush", "title": "x"}),
                    content_type="application/json")
    assert r.status_code == 403


def test_下書きは送らず記録だけ(client, api, fake_onesignal):
    答 = _post(client, {"type": "broadcastPush", "mode": "draft", "title": "秋の教室", "body": "募集中"})
    assert 答["status"] == "ok" and 答["rowIdx"] == 2
    p = PushNotice.objects.get(sheet_row=2)
    assert p.status == "下書き" and p.sent_at is None
    assert fake_onesignal == []


def test_送ると全員宛てで記録に通知IDが付く(client, api, fake_onesignal):
    答 = _post(client, {"type": "broadcastPush", "mode": "send", "title": "秋の教室", "body": "募集中",
                        "targetStatus": "all", "targetPage": "news"})
    assert 答["status"] == "ok", 答
    body = fake_onesignal[0]
    assert "filters" in body and "include_subscription_ids" not in body
    assert body["data"]["openPage"] == "news"
    p = PushNotice.objects.get(sheet_row=2)
    assert p.status == "送信済み" and p.notification_id == "n-777" and p.recipient_count == 97
    assert p.result == "送信済み" and p.sent_at is not None


def test_テスト送信は指定の端末だけ(client, api, fake_onesignal):
    users = [{"memberId": "MYM-0001", "name": "山田", "subscription": "sub-abc"}]
    答 = _post(client, {"type": "broadcastPush", "mode": "test", "title": "試し", "body": "本文",
                        "targetUsers": users, "targetDetail": "テスト送信:山田", "recipientCount": 1})
    assert 答["status"] == "ok", 答
    assert fake_onesignal[0]["include_subscription_ids"] == ["sub-abc"]
    assert PushNotice.objects.get().result == "テスト送信済み"


def test_自動送信は投稿と同じ内容で全員へ(client, api, fake_onesignal):
    答 = _post(client, {"type": "broadcastPush", "mode": "auto", "title": "📝 新しい投稿",
                        "body": "NEWSが更新されました", "targetPage": "news"})
    assert 答["status"] == "ok" and 答["sent"] is True
    assert PushNotice.objects.get().status == "自動送信済み"


def test_送れなければ送信失敗と理由が残る(client, api, monkeypatch):
    def boom(req, timeout=0):
        raise onesignal.urllib.error.URLError("down")

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", boom)
    答 = _post(client, {"type": "broadcastPush", "mode": "send", "title": "x", "body": "y"})
    assert 答["status"] == "error"
    p = PushNotice.objects.get()
    assert p.status == "送信失敗" and "down" in p.result


def test_鍵が無ければ送れずに失敗が残る(client, settings, api):
    settings.ONESIGNAL_REST_API_KEY = ""
    答 = _post(client, {"type": "broadcastPush", "mode": "send", "title": "x", "body": "y"})
    assert 答["status"] == "error"
    assert PushNotice.objects.get().status == "送信失敗"


def test_予約分は時刻が来たものだけ送る(client, api, fake_onesignal):
    now = timezone.localtime()
    past = (now - timedelta(minutes=10)).strftime("%Y-%m-%dT%H:%M")
    future = (now + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M")
    _post(client, {"type": "broadcastPush", "mode": "schedule", "title": "そろそろ", "scheduledAt": past})
    _post(client, {"type": "broadcastPush", "mode": "schedule", "title": "まだ", "scheduledAt": future})
    答 = _post(client, {"type": "processScheduledPushQueue"})
    assert 答 == {"status": "ok", "processed": 1}
    assert len(fake_onesignal) == 1
    assert PushNotice.objects.get(title="そろそろ").result == "予約送信済み"
    assert PushNotice.objects.get(title="まだ").status == "予約済み"
    # もう一度呼んでも二度は送らない
    _post(client, {"type": "processScheduledPushQueue"})
    assert len(fake_onesignal) == 1


def test_予約に日時が無ければ断る(client, api):
    答 = _post(client, {"type": "broadcastPush", "mode": "schedule", "title": "x"})
    assert 答["status"] == "error"


def test_rowIdxがあれば同じ行を上書きする(client, api, fake_onesignal):
    答 = _post(client, {"type": "broadcastPush", "mode": "draft", "title": "下書き"})
    row = 答["rowIdx"]
    _post(client, {"type": "broadcastPush", "mode": "send", "rowIdx": row, "title": "本番"})
    assert PushNotice.objects.count() == 1
    assert PushNotice.objects.get().title == "本番"


def test_履歴の削除は印だけで旧管理アプリの一覧から消える(client, api):
    _post(client, {"type": "broadcastPush", "mode": "draft", "title": "消す"})
    答 = _post(client, {"type": "deleteRow", "sheet": "PUSH", "rowIdx": 2})
    assert 答["status"] == "ok"
    p = PushNotice.objects.get(sheet_row=2)
    assert p.deleted is True and p.deleted_at is not None
    一覧 = client.get("/api?action=getPushNotices").json()
    assert 一覧["notices"] == []


def test_まとめて削除(client, api):
    for t in ("a", "b", "c"):
        _post(client, {"type": "broadcastPush", "mode": "draft", "title": t})
    答 = _post(client, {"type": "deleteRows", "sheet": "PUSH", "rowIdxs": [2, 3]})
    assert 答 == {"status": "ok", "deleted": 2}
    assert PushNotice.objects.filter(deleted=False).count() == 1


def test_お知らせの削除は通知の表に触らない(client, api):
    _post(client, {"type": "broadcastPush", "mode": "draft", "title": "残る"})
    答 = _post(client, {"type": "deleteRow", "sheet": "BLOG", "rowIdx": 2})
    assert 答["status"] == "error"  # お知らせの行2は無い
    assert PushNotice.objects.get().deleted is False


# ========== 管理画面 ==========


def test_管理画面から全員へ送れる(as_owner, api, fake_onesignal):
    r = as_owner.post("/manage/push/send/", {"mode": "send", "title": "教室のご案内", "body": "本文", "targetPage": "calendar"})
    assert r.status_code == 302
    assert fake_onesignal[0]["data"]["openPage"] == "calendar"
    page = as_owner.get("/manage/push/").content.decode()
    assert "教室のご案内" in page and "送信済み" in page


def test_管理画面の予約と削除(as_owner, api, fake_onesignal):
    when = (timezone.localtime() + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M")
    as_owner.post("/manage/push/send/", {"mode": "schedule", "title": "明日", "scheduledAt": when})
    p = PushNotice.objects.get()
    assert p.status == "予約済み" and fake_onesignal == []
    as_owner.post(f"/manage/push/{p.sheet_row}/delete/")
    p.refresh_from_db()
    assert p.deleted is True


def test_鍵が無いと管理画面に注意が出る(as_owner, settings):
    settings.ONESIGNAL_APP_ID = ""
    assert "設定（OneSignal の鍵）がサーバーにありません" in as_owner.get("/manage/push/").content.decode()
