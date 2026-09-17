"""ビジリス管理「お客様の画面」（apps/bijiris/views_preview.py）。

院長の依頼（2026-09-17）: お客様のアプリ画面を、いま作っている方（develop・手元のファイル）と
お客様に出ている方（main・公開先）で見比べられるようにする。
"""

import pytest

from apps.bijiris.views_preview import 公開先

pytestmark = pytest.mark.django_db


@pytest.fixture
def 置き場(tmp_path, settings):
    """手元の bijiris/customer-app の代わり。index.html と、下のファイルを1つ置く。"""
    app = tmp_path / "customer-app"
    (app / "icons").mkdir(parents=True)
    (app / "index.html").write_text("<h1>まゆみ助産院 アンケート</h1>", encoding="utf-8")
    (app / "styles.css").write_text("body{color:#333}", encoding="utf-8")
    (tmp_path / "秘密.txt").write_text("外にあるもの", encoding="utf-8")
    settings.BIJIRIS_DIR = str(tmp_path)
    return app


def test_タブが出る(as_owner, 置き場):
    page = as_owner.get("/manage/bijiris/preview/")
    assert page.status_code == 200
    html = page.content.decode()
    i = html.find('<nav class="page-tabs"')
    tabs = html[i:html.find("</nav>", i)]
    # 6画面の並びの最後に「お客様の画面」。開いているのはこのタブ
    for t in ["集計", "アンケート管理", "回答管理", "顧客管理", "特典", "回数券分析", "お客様の画面"]:
        assert t in tabs, t
    assert 'class="active">お客様の画面</a>' in tabs
    # 顧客管理と取り違えていない（url_name の頭かぶり）
    assert 'class="active">顧客管理</a>' not in tabs
    assert 'class="active">お客様の画面</a>' not in as_owner.get("/manage/bijiris/customers/").content.decode()


def test_切り替えが効く(as_owner, 置き場):
    # 既定は「いま作っている方」。枠の中は手元のファイルの配り口
    html = as_owner.get("/manage/bijiris/preview/").content.decode()
    assert "/manage/bijiris/preview/files/" in html and 公開先 not in html
    # 「お客様に出ている方」は公開先の住所がそのまま埋まっている
    html = as_owner.get("/manage/bijiris/preview/?which=live").content.decode()
    assert 公開先 in html and "/manage/bijiris/preview/files/" not in html
    # 幅は 375px（スマホ）とパソコン（枠いっぱい）
    assert "width:375px" in as_owner.get("/manage/bijiris/preview/?w=375").content.decode()
    assert "width:100%" in as_owner.get("/manage/bijiris/preview/?w=0").content.decode()
    # 知らない値は既定に戻す（落ちない）
    html = as_owner.get("/manage/bijiris/preview/?which=よそ&w=9999").content.decode()
    assert "width:375px" in html and "/manage/bijiris/preview/files/" in html
    # 「別の窓で開く」がどちらにも付いている
    assert "別の窓で開く" in html and "別の窓で開く" in as_owner.get("/manage/bijiris/preview/?which=live").content.decode()


def test_手元のファイルを配り枠に出せる(as_owner, 置き場):
    r = as_owner.get("/manage/bijiris/preview/files/")
    assert r.status_code == 200 and b"\xe3\x82\xa2\xe3\x83\xb3\xe3\x82\xb1\xe3\x83\xbc\xe3\x83\x88" in b"".join(r.streaming_content)
    # 枠（iframe）に出せること。既定の DENY だと「接続が拒否されました」になる
    assert r.get("X-Frame-Options", "") != "DENY"
    assert as_owner.get("/manage/bijiris/preview/files/styles.css")["Content-Type"] == "text/css"
    assert as_owner.get("/manage/bijiris/preview/files/nothing.js").status_code == 404


def test_上の道は弾く(as_owner, 置き場):
    """customer-app の外は出さない。"""
    for path in ["../秘密.txt", "icons/../../秘密.txt", "..%2F秘密.txt", "a/../../秘密.txt"]:
        assert as_owner.get("/manage/bijiris/preview/files/" + path).status_code == 404, path


def test_置き場所が無くても画面が開く(as_owner, settings, tmp_path):
    settings.BIJIRIS_DIR = str(tmp_path / "ない")
    page = as_owner.get("/manage/bijiris/preview/")
    assert page.status_code == 200
    html = page.content.decode()
    # 落ちずに、手元が無いことを伝え、公開の方は見られる
    assert "手元にファイルがありません" in html and 公開先 in html
    assert as_owner.get("/manage/bijiris/preview/files/").status_code == 404
    # 設定そのものが空でも落ちない
    settings.BIJIRIS_DIR = ""
    assert as_owner.get("/manage/bijiris/preview/").status_code == 200


def test_スタッフは403で未ログインはログインへ(client, staff):
    for url in ["/manage/bijiris/preview/", "/manage/bijiris/preview/files/"]:
        r = client.get(url)
        assert r.status_code == 302 and r["Location"].startswith("/manage/login/"), url
    client.force_login(staff)
    for url in ["/manage/bijiris/preview/", "/manage/bijiris/preview/files/"]:
        assert client.get(url).status_code == 403, url
