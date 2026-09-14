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
