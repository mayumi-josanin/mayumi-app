"""QRコード案内。**中身は旧管理アプリ（admin/index.html の #page-qrcode）と同じ**
（アプリ配布用QR・来院スタンプ用QR・配布用資料）。

**QRの絵は外の作り手（api.qrserver.com）に頼らない。**院内から外へ出られないと出ないため、
このサーバーが描く。ここではその道（/manage/qrcode/image/）が絵を返すことを見る。
"""

import pytest

from apps.manage.views_qrcode import アプリの住所, スタンプの住所, 住所を決める
from apps.records.models import AppSetting

pytestmark = pytest.mark.django_db

既定 = "https://mayumi-josanin.github.io/mayumi-app/"


def test_まゆみ以外は入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/qrcode/").status_code == 403
    assert client.get("/manage/qrcode/image/?kind=app").status_code == 403


def test_住所の既定が入っている():
    """移行 0011 で、いま配っている住所が最初から入っている。"""
    assert アプリの住所() == 既定


def test_2つのQRが出る(as_owner):
    page = as_owner.get("/manage/qrcode/").content.decode()
    assert "アプリ配布用QRコード" in page
    assert "来院スタンプ用QRコード" in page
    # 絵は外の作り手ではなく、このサーバーの道を見ている
    assert "/manage/qrcode/image/?kind=app" in page
    assert "/manage/qrcode/image/?kind=stamp" in page
    assert "api.qrserver.com" not in page
    # スキャン用URL（アプリの住所と、スタンプ用の中継ページの住所）
    assert 既定 in page
    assert "stamp-launch.html?action=add_stamp" in page


def test_スタンプの説明書き4行がそのまま出る(as_owner):
    page = as_owner.get("/manage/qrcode/").content.decode()
    for 行 in [
        "iPhone / Android の標準カメラで読み取ると、中継ページが開いてアプリ起動だけを試します。",
        "mayumijosanin://stamp?open=home",
        "来院スタンプは、アプリを開いた後にアプリ内の「カメラを起動して読み取る」から同じQRコードを読み取った時だけ付与されます。",
        "iPhone標準カメラでは「使用可能なデータがみつかりません」になるため使わないでください。",
    ]:
        assert 行 in page, 行


def test_配布用資料が同じ画面にあり印刷用の見え方がある(as_owner):
    """旧アプリは別の窓を開いていた。窓はブラウザに止められることがあるので同じ画面に置く。"""
    page = as_owner.get("/manage/qrcode/").content.decode()
    assert "配布用資料" in page
    assert "@media print" in page
    assert "window.print()" in page
    # 旧アプリの資料の文言がそのまま入っている
    assert "公式アプリのご案内" in page
    assert "スマートフォンでスキャンするだけで、" in page
    assert "A4またはB5サイズでの印刷を推奨します。" in page
    assert "まゆみ助産院 公式アプリ" in page


def test_QRの絵が返る(as_owner):
    for kind in ("app", "stamp"):
        res = as_owner.get(f"/manage/qrcode/image/?kind={kind}")
        assert res.status_code == 200
        assert res["Content-Type"] == "image/png"
        assert res.content.startswith(b"\x89PNG")
        assert len(res.content) > 200
    res = as_owner.get("/manage/qrcode/image/?kind=app&format=svg")
    assert res.status_code == 200
    assert res["Content-Type"] == "image/svg+xml"
    assert b"<svg" in res.content


def test_画像で保存は保存になる(as_owner):
    res = as_owner.get("/manage/qrcode/image/?kind=app&download=1")
    assert res.status_code == 200
    assert res["Content-Disposition"].startswith("attachment;")


def test_住所を変えるとQRの中身も変わる(as_owner):
    元 = as_owner.get("/manage/qrcode/image/?kind=app").content
    住所を決める("https://example.com/app/")
    後 = as_owner.get("/manage/qrcode/image/?kind=app").content
    assert 元 != 後
    page = as_owner.get("/manage/qrcode/").content.decode()
    assert "https://example.com/app/" in page
    assert "https://example.com/app/stamp-launch.html?action=add_stamp" in page


def test_住所が決まっていないときは案内が出る(as_owner):
    AppSetting.objects.filter(pk="app_public_url").delete()
    page = as_owner.get("/manage/qrcode/").content.decode()
    assert "お客様アプリの住所が決まっていないため、QRコードを作れません。" in page
    # 決める場所への行き先が付いている
    assert "/manage/system/#app-url" in page
    # 絵は描かない（読んでも何も起きない紙が出回らないように）
    assert as_owner.get("/manage/qrcode/image/?kind=app").status_code == 404
    assert as_owner.get("/manage/qrcode/image/?kind=stamp").status_code == 404


def test_知らない種類は404(as_owner):
    assert as_owner.get("/manage/qrcode/image/?kind=なにか").status_code == 404


def test_スタンプの住所の組み立て():
    assert スタンプの住所("https://example.com/app/") == "https://example.com/app/stamp-launch.html?action=add_stamp"
    # 形になっていない文字のときは空（旧アプリの buildStampLaunchUrl と同じ）
    assert スタンプの住所("") == ""
    assert スタンプの住所("ただの文字") == ""


def test_システム管理から住所を決められる(as_owner):
    res = as_owner.post("/manage/system/app-url/", {"app_url": "https://example.com/app/"})
    assert res.status_code == 302
    assert アプリの住所() == "https://example.com/app/"
    page = as_owner.get("/manage/system/").content.decode()
    assert "お客様アプリの住所" in page and "https://example.com/app/" in page
    # https:// から始まらないものは断る（QRを読んでも開けない住所を保存しない）
    res = as_owner.post("/manage/system/app-url/", {"app_url": "example.com"})
    assert アプリの住所() == "https://example.com/app/"


def test_左メニューのアプリ管理に出る(as_owner):
    page = as_owner.get("/manage/qrcode/").content.decode()
    tabs = page[page.find('<nav class="page-tabs"'):]
    tabs = tabs[:tabs.find("</nav>")]
    assert 'class="active">QRコード案内</a>' in tabs
    # まとまりは「設定」（旧アプリと同じく並びの最後）
    side = page[page.find('<nav class="sidebar-nav">'):page.find("</nav>")]
    assert 'class="active"><span class="nav-icon">🛠️</span> 設定</a>' in side
