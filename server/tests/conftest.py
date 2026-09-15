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


# ---------------------------------------------------------------------------
# 会員の窓口（test_gasapi_member*.py）向け
# ---------------------------------------------------------------------------

@pytest.fixture
def 会員試験の下ごしらえ(settings):
    """合鍵を入れ、回数制限の数え上げ（cache）を空にし、ハッシュを軽くする。

    ・合鍵: 管理者向けの窓口は X-Api-Key が要る。空だと 503 で全部止まる。
    ・cache: ログイン・復元の「10回はずしたらお休み」は cache で数えている。
      試験ごとに空にしないと、前の試験のはずれが次の試験に持ち越される。
    ・ハッシュ: 既定の PBKDF2 は1回 0.1 秒かかる。会員を何十人も作る試験で
      積み上がるので、試験では MD5 にする。照合の作法（check_password）は同じ。
    """
    from django.core.cache import cache

    from tests.gasapi_member_support import 合鍵

    settings.API_KEY = 合鍵
    settings.PASSWORD_HASHERS = ["django.contrib.auth.hashers.MD5PasswordHasher"]
    cache.clear()
    yield
    cache.clear()
