"""カレンダー管理。読み書きは gasapi/admin_calendar（GAS の転送先）と同じ。"""

import io
import json
from datetime import date

import pytest
from django.utils import timezone

from apps.content import onesignal
from apps.content.models import CalendarEvent, Menu, PushNotice
from apps.gasapi.admin_calendar import 一覧

pytestmark = pytest.mark.django_db


@pytest.fixture
def fake_onesignal(monkeypatch, settings):
    settings.ONESIGNAL_APP_ID = "app"
    settings.ONESIGNAL_REST_API_KEY = "key"
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append(json.loads(req.data))
        return _Res(json.dumps({"id": "n-1", "recipients": 5}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


def _form(**extra):
    d = {"date": "2026-10-03", "title": "ベビーマッサージ教室", "category": "教室", "menuRowIdx": "",
         "desc": "10:00〜", "color": "#f48fb1", "publishAt": "", "linkUrl": "", "linkButtonText": "",
         "notice_listed": "on", "status": "公開"}
    d.update(extra)
    return d


def test_日付を複数選ぶと選んだ数だけ作る(as_owner):
    r = as_owner.post("/manage/calendar/new/", {**_form(), "dates": ["2026-10-03", "2026-10-10", "2026-10-17"]})
    assert r.status_code == 302, r.content.decode()[:500]
    rows = list(CalendarEvent.objects.order_by("sheet_row"))
    assert [c.sheet_row for c in rows] == [2, 3, 4]
    assert [c.event_on for c in rows] == [date(2026, 10, 3), date(2026, 10, 10), date(2026, 10, 17)]
    assert all(c.title == "ベビーマッサージ教室" and c.color == "#f48fb1" and c.published for c in rows)
    assert len(一覧()["events"]) == 3


def test_追加した日が無ければ日付1つで作る(as_owner):
    as_owner.post("/manage/calendar/new/", _form())
    assert CalendarEvent.objects.count() == 1


def test_日付が無ければ断る(as_owner):
    r = as_owner.post("/manage/calendar/new/", _form(date=""))
    assert r.status_code == 200
    assert CalendarEvent.objects.count() == 0


def test_対象メニューと休診(as_owner):
    Menu.objects.create(sheet_row=5, name="ベビマ", published=True)
    as_owner.post("/manage/calendar/new/", _form(menuRowIdx="5"))
    as_owner.post("/manage/calendar/new/", _form(title="休診日", category="休診", date="2026-10-04", color="#e57373"))
    a = CalendarEvent.objects.get(title="ベビーマッサージ教室")
    assert a.menu_row == 5
    rows = {e["rowIdx"]: e for e in 一覧()["events"]}
    assert rows[a.sheet_row]["menuRowIdx"] == 5
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert "ベビマ" in page and "badge-red" in page


def test_一覧は年月で絞れる(as_owner):
    as_owner.post("/manage/calendar/new/", _form(date="2026-10-03"))
    as_owner.post("/manage/calendar/new/", _form(title="11月の予定", date="2026-11-03"))
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert "ベビーマッサージ教室" in page and "11月の予定" not in page
    page = as_owner.get("/manage/calendar/?year=2026").content.decode()
    assert "11月の予定" in page


def test_修正は日付1つだけを直す(as_owner):
    as_owner.post("/manage/calendar/new/", {**_form(), "dates": ["2026-10-03", "2026-10-10"]})
    c = CalendarEvent.objects.get(event_on=date(2026, 10, 10))
    as_owner.post(f"/manage/calendar/{c.sheet_row}/", _form(date="2026-10-11", title="改", dates=["2026-12-01"]))
    c.refresh_from_db()
    assert c.event_on == date(2026, 10, 11) and c.title == "改"
    assert CalendarEvent.objects.count() == 2  # dates は修正では無視する


def test_公開切替と削除は印だけ(as_owner):
    as_owner.post("/manage/calendar/new/", _form())
    c = CalendarEvent.objects.get()
    as_owner.post(f"/manage/calendar/{c.sheet_row}/toggle/")
    c.refresh_from_db()
    assert c.published is False
    as_owner.post(f"/manage/calendar/{c.sheet_row}/delete/")
    c.refresh_from_db()
    assert c.deleted is True and CalendarEvent.objects.filter(pk=c.pk).exists()
    assert 一覧()["events"] == []


def test_通知の文言と条件(as_owner, fake_onesignal):
    as_owner.post("/manage/calendar/new/", _form(send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "📅 ベビーマッサージ教室"
    assert fake_onesignal[-1]["contents"]["ja"] == "カレンダーが更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "calendar"
    assert PushNotice.objects.count() == 1
    as_owner.post("/manage/calendar/new/", _form(title="下書き", status="非公開", send_push="on"))
    assert len(fake_onesignal) == 1


def test_画像はカレンダーの置き場へ(as_owner):
    from django.core.files.uploadedfile import SimpleUploadedFile
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (400, 300), (10, 20, 30)).save(buf, format="PNG")
    as_owner.post("/manage/calendar/new/", {**_form(), "images": [SimpleUploadedFile("a.png", buf.getvalue(), content_type="image/png")]})
    c = CalendarEvent.objects.get()
    assert "/media/calendar/" in c.image_url
    assert as_owner.get("/media/calendar/nothing.jpg").status_code == 404
