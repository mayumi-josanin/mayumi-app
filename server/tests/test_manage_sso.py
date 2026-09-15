"""予約管理（mayumi-reserve）とのログインの受け渡し。"""

import time
from urllib.parse import parse_qs, urlparse

import pytest

from apps.manage import sso

pytestmark = pytest.mark.django_db

SECRET = "shared-secret-for-test"


@pytest.fixture
def keys(settings):
    settings.MANAGE_SSO_SECRET = SECRET
    settings.RESERVE_MANAGE_URL = "https://reserve.example.com:10001"


def test_札は署名と期限と使い捨ての印を持つ():
    札 = sso.札を作る("まゆみ", SECRET)
    中 = sso.札を確かめる(札, SECRET)
    assert 中 and 中["role"] == "まゆみ" and 中["nonce"]
    assert sso.札を確かめる(札, "別の合鍵") is None
    assert sso.札を確かめる(札 + "x", SECRET) is None


def test_期限切れの札は通らない(monkeypatch):
    札 = sso.札を作る("まゆみ", SECRET)
    now = time.time()
    monkeypatch.setattr(time, "time", lambda: now + 10_000)
    assert sso.札を確かめる(札, SECRET) is None


def test_予約管理へは札を付けて飛ぶ(as_owner, keys):
    r = as_owner.get("/manage/go/reserve/?to=/manage/customers/")
    assert r.status_code == 302
    u = urlparse(r["Location"])
    assert u.netloc == "reserve.example.com:10001" and u.path == "/manage/sso/"
    q = parse_qs(u.query)
    assert q["next"] == ["/manage/customers/"]
    assert sso.札を確かめる(q["t"][0], SECRET)["role"] == "まゆみ"


def test_外の住所へは飛ばさない(as_owner, keys):
    r = as_owner.get("/manage/go/reserve/?to=https://evil.example.com/")
    assert parse_qs(urlparse(r["Location"]).query)["next"] == ["/manage/"]


def test_入っていなければまずログイン(client, keys):
    r = client.get("/manage/go/reserve/?to=/manage/")
    assert r["Location"].startswith("/manage/login/")


def test_合鍵が無ければただ相手へ(as_owner, settings):
    settings.MANAGE_SSO_SECRET = ""
    settings.RESERVE_MANAGE_URL = "https://reserve.example.com:10001"
    r = as_owner.get("/manage/go/reserve/?to=/manage/")
    assert r["Location"] == "https://reserve.example.com:10001/manage/"


def test_予約管理からの札で入れる(client, owner, keys):
    札 = sso.札を作る("まゆみ", SECRET)
    r = client.get(f"/manage/sso/?t={札}&next=/manage/orders/")
    assert r.status_code == 302 and r["Location"] == "/manage/orders/"
    assert client.get("/manage/orders/").status_code == 200
    # 同じ札は二度使えない
    other = client.__class__()
    r2 = other.get(f"/manage/sso/?t={札}&next=/manage/orders/")
    assert r2["Location"].startswith("/manage/login/")


def test_偽の札では入れない(client, owner, keys):
    r = client.get("/manage/sso/?t=abc.def&next=/manage/")
    assert r["Location"].startswith("/manage/login/")
    assert client.get("/manage/news/").status_code == 302


def test_その役割の人がいなければログインへ(client, keys):
    札 = sso.札を作る("スタッフ", SECRET)
    r = client.get(f"/manage/sso/?t={札}&next=/manage/")
    assert r["Location"].startswith("/manage/login/")
