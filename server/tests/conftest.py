import pytest
from django.contrib.auth.models import Group, User

from apps.manage.permissions import OWNER, STAFF, ensure_groups


@pytest.fixture(autouse=True)
def _plain_static(settings, tmp_path):
    settings.STORAGES = {
        "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
        "staticfiles": {"BACKEND": "django.contrib.staticfiles.storage.StaticFilesStorage"},
    }
    settings.MEDIA_ROOT = tmp_path / "media"
    settings.PUBLIC_BASE_URL = "https://api.example.com"


@pytest.fixture
def owner(db):
    ensure_groups()
    user = User.objects.create_user(username=OWNER, password="0714")
    user.groups.add(Group.objects.get(name=OWNER))
    return user


@pytest.fixture
def staff(db):
    ensure_groups()
    user = User.objects.create_user(username=STAFF, password="1234")
    user.groups.add(Group.objects.get(name=STAFF))
    return user


@pytest.fixture
def as_owner(client, owner):
    client.force_login(owner)
    return client


import io  # noqa: E402
import json  # noqa: E402

from apps.content import onesignal  # noqa: E402


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
