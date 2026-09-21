"""前の管理アプリを閉じたこと（2026-09-21）。

画面13枚と、画面以外の設定3つを全部 Django の管理画面へ移し終えた。
前のアプリは「移転しました」の画面に差し替え、左メニューのリンクも外した。

**差し替えだけでは足りない。**前のアプリは画面を端末に覚えさせる仕掛けを持っていて、
そのままだと古い画面が出続ける。仕掛け自身も、自分を外すものに差し替えてある。
"""

from pathlib import Path

import pytest

リポジトリ = Path(__file__).resolve().parent.parent.parent

pytestmark = pytest.mark.django_db


def test_前のアプリは移転しましたの画面になっている():
    s = (リポジトリ / "admin" / "index.html").read_text(encoding="utf-8")
    assert "管理アプリは新しい管理画面へ移りました" in s
    assert len(s) < 8000, "前のアプリの中身がまだ残っています"


def test_覚えさせる仕掛けは自分を外す():
    s = (リポジトリ / "admin" / "admin-sw.js").read_text(encoding="utf-8")
    assert "unregister" in s
    # お客様アプリのものまで消さない
    assert "mayumi-admin-" in s
    assert "ADMIN_SHELL_ASSETS" not in s, "前の仕掛けが残っています"


def test_左メニューに前のアプリへのリンクが無い(as_owner):
    page = as_owner.get("/manage/members/").content.decode()
    assert "mayumi-app/admin/" not in page
