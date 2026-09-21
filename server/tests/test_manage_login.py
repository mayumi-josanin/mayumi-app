"""管理画面のログイン。役割を選んでパスコードで入る。"""

import pytest

pytestmark = pytest.mark.django_db


def test_入っていなければログインへ(client):
    r = client.get("/manage/news/")
    assert r.status_code == 302
    assert r["Location"].startswith("/manage/login/")


def test_役割とパスコードで入れる(client, owner):
    r = client.post("/manage/login/", {"role": "まゆみ", "passcode": "0714"})
    assert r.status_code == 302
    assert r["Location"] == "/manage/news/"
    assert client.get("/manage/news/").status_code == 200


def test_全角のパスコードでも入れる(client, owner):
    r = client.post("/manage/login/", {"role": "まゆみ", "passcode": "０７１４"})
    assert r.status_code == 302


def test_違う役割のパスコードでは入れない(client, owner, staff):
    r = client.post("/manage/login/", {"role": "まゆみ", "passcode": "1234"})  # スタッフのパスコード
    assert r.status_code == 200
    assert "違うようです" in r.content.decode()


def test_外の住所へは飛ばさない(client, owner):
    r = client.post(
        "/manage/login/",
        {"role": "まゆみ", "passcode": "0714", "next": "https://evil.example.com/"},
    )
    assert r["Location"] == "/manage/news/"


def test_間違いが続くと待たせる(client, owner):
    for _ in range(5):
        client.post("/manage/login/", {"role": "まゆみ", "passcode": "9999"})
    r = client.post("/manage/login/", {"role": "まゆみ", "passcode": "0714"})
    assert r.status_code == 429


def test_ログアウトはPOSTだけ(as_owner):
    assert as_owner.get("/manage/logout/").status_code == 405
    r = as_owner.post("/manage/logout/")
    assert r.status_code == 302
    assert as_owner.get("/manage/news/").status_code == 302


def test_パスコードは環境変数から決める(db, monkeypatch):
    from django.core.management import call_command

    monkeypatch.setenv("MANAGE_PASSCODE", "２０２６")  # 全角でも通す
    call_command("manage_passcode", "owner")
    from django.contrib.auth.models import User

    user = User.objects.get(username="まゆみ")
    assert user.check_password("2026")
    assert user.groups.filter(name="まゆみ").exists()


def test_スタッフは予約管理だけ(client, staff):
    """院長の決定（2026-09-16）: スタッフがログインしたら予約管理だけ。ログインは通し、そのまま予約管理へ受け渡す。"""
    r = client.post("/manage/login/", {"role": "スタッフ", "passcode": "1234"})
    assert r.status_code == 302 and r["Location"] == "/manage/go/reserve/?to=/manage/"
    # アプリ管理の画面は開けない
    assert client.get("/manage/orders/").status_code == 403
    assert client.get("/manage/news/").status_code == 403
    assert client.get("/manage/dev/").status_code == 403
    # 入口も予約管理へ
    assert client.get("/manage/")["Location"] == "/manage/go/reserve/?to=/manage/"


def test_スタッフには左メニューの予約管理だけ(rf, staff, owner):
    from apps.manage import navigation

    req = rf.get("/manage/"); req.user = staff; req.resolver_match = None
    sections = navigation.組み立てる(req)["nav_sections"]
    assert [s["label"] for s in sections] == ["予約管理"]
    req.user = owner
    assert [s["label"] for s in navigation.組み立てる(req)["nav_sections"]] == [
        "アプリ管理", "公式サイト", "ビジリス", "経費管理", "お客様の画面", "開発", "予約管理"]


def test_スタッフの札では受け渡しも断る(client, staff, settings):
    from apps.manage import sso

    settings.MANAGE_SSO_SECRET = "s"
    r = client.get(f"/manage/sso/?t={sso.札を作る('スタッフ', 's')}&next=/manage/orders/")
    assert r["Location"].startswith("/manage/login/")
    assert client.get("/manage/orders/").status_code == 302


def test_CSSとJSには版が付く(as_owner):
    """端末に古い CSS / JS が残って、直したものが効かないことを防ぐ（2026-09-16）。"""
    page = as_owner.get("/manage/orders/").content.decode()
    assert "manage/style.css?v=" in page and "manage/app.js?v=" in page
