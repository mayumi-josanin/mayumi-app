"""お客様アプリへの通知（OneSignal）を、この管理画面から送る。

GAS の sendAutoPush と同じ内容（タグで全員・開くページ）。
**設定が無ければ送らない**、送れなくても投稿は残る。
"""

import io
import json

import pytest
from django.utils import timezone

from apps.content.models import Category, News, PushNotice
from apps.content import onesignal
from apps.manage import push  # noqa: F401

pytestmark = pytest.mark.django_db


@pytest.fixture
def keys(settings):
    settings.ONESIGNAL_APP_ID = "app-0001"
    settings.ONESIGNAL_REST_API_KEY = "key-0001"


@pytest.fixture
def fake_onesignal(monkeypatch):
    """OneSignal へ実際には行かない。送った中身を控える。"""
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append({"url": req.full_url, "headers": dict(req.header_items()), "body": json.loads(req.data)})
        return _Res(json.dumps({"id": "n-123", "recipients": 42}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


def _form(**extra):
    Category.objects.get_or_create(sheet_row=2, name="イベント情報", defaults={"kind": "お知らせ"})
    data = {
        "posted_on": "2026-09-14", "title": "マッサージ教室", "category": "イベント情報", "icon": "📢",
        "body": "本文", "publish_at": "", "link_url": "", "button_text": "", "notice_listed": "on",
        "status": "公開",
    }
    data.update(extra)
    return data


def test_設定が無ければ送らず投稿は残る(as_owner, settings):
    settings.ONESIGNAL_APP_ID = ""
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert News.objects.filter(title="マッサージ教室").exists()
    assert PushNotice.objects.count() == 0


def test_チェックすると全員へ送り記録が残る(as_owner, keys, fake_onesignal):
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert len(fake_onesignal) == 1
    body = fake_onesignal[0]["body"]
    assert body["app_id"] == "app-0001"
    assert body["headings"]["ja"] == "📝 マッサージ教室"
    assert body["contents"]["ja"] == "NEWSが更新されました"
    assert body["filters"] == [{"field": "tag", "key": "app_scope", "relation": "=", "value": "mayumi_josanin_app"}]
    assert body["url"].endswith("?open=news") and body["data"]["openPage"] == "news"
    assert fake_onesignal[0]["headers"]["Authorization"] == "Basic key-0001"

    rec = PushNotice.objects.get()
    assert rec.status == "自動送信済み"
    assert rec.notification_id == "n-123"
    assert rec.recipient_count == 42
    assert rec.sent_at is not None
    assert rec.sheet_row == 2


def test_チェックしなければ送らない(as_owner, keys, fake_onesignal):
    as_owner.post("/manage/news/new/", _form())
    assert fake_onesignal == []


def test_下書きや公開開始が先なら送らない(as_owner, keys, fake_onesignal):
    as_owner.post("/manage/news/new/", _form(send_push="on", status="非公開"))
    future = (timezone.localtime() + timezone.timedelta(days=2)).strftime("%Y-%m-%dT%H:%M")
    as_owner.post("/manage/news/new/", _form(send_push="on", title="先の予定", publish_at=future))
    assert fake_onesignal == []


def test_送れなくても投稿は残り失敗が記録される(as_owner, keys, monkeypatch):
    def boom(req, timeout=0):
        raise onesignal.urllib.error.URLError("down")

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", boom)
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert News.objects.filter(title="マッサージ教室").exists()
    rec = PushNotice.objects.get()
    assert rec.status == "送信失敗"
