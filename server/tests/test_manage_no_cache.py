"""管理画面の画面は、ブラウザに覚えさせない（2026-09-20）。

直してサーバーに入れたのに、院長の画面では前のままだった。
直しに入れた赤字の案内すら出なかったので、古い画面がそのまま使われていたと分かった。
"""

import pytest

pytestmark = pytest.mark.django_db


def test_管理画面は覚えさせない(as_owner):
    r = as_owner.get("/manage/menus/")
    assert r.status_code == 200
    assert "no-store" in r.headers.get("Cache-Control", "")


def test_ログインの画面も覚えさせない(client):
    r = client.get("/manage/login/")
    assert "no-store" in r.headers.get("Cache-Control", "")


def test_書き出したものには付けない(as_owner):
    """CSV などは画面ではないので触らない。"""
    r = as_owner.get("/manage/members/")
    assert "text/html" in r.headers.get("Content-Type", "")
    assert "no-store" in r.headers.get("Cache-Control", "")
