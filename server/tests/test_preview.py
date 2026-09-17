"""お客様の画面（apps/preview）。

院長の依頼（2026-09-18）:「お客様アプリがあるやつ（ビジリスアプリ・まゆみ助産院お客様アプリ・予約システム）を
develop 環境で確認して見れるようにできますか」。3つを1か所にまとめ、
「いま作っている方（develop）」と「お客様に出ている方（main）」を見比べられるようにした。

**いちばん大事なのは、配ってはいけないものが配られないこと。**server/（この Django のコード・.env・backups）は
写しの中に入っているので、通してよい一覧で確かに止まることを試験する。
"""

import pytest

from apps.preview.views import 種類たち

pytestmark = pytest.mark.django_db

アプリ公開先 = 種類たち["app"]["公開先"]
ビジリス公開先 = 種類たち["bijiris"]["公開先"]


@pytest.fixture
def 写し(tmp_path, settings):
    """develop の写しの代わり。お客様アプリ本体と、ビジリスと、**配ってはいけないもの**を置く。"""
    root = tmp_path
    (root / "icons").mkdir()
    (root / "index.html").write_text("<h1>まゆみ助産院 お客様アプリ</h1>", encoding="utf-8")
    (root / "style.css").write_text("body{color:#333}", encoding="utf-8")
    (root / "icons" / "icon-192.png").write_bytes(b"png")
    # ビジリスのお客様アプリ（同じ写しの下）
    bij = root / "bijiris" / "customer-app"
    (bij / "icons").mkdir(parents=True)
    (bij / "index.html").write_text("<h1>まゆみ助産院 アンケート</h1>", encoding="utf-8")
    (bij / "styles.css").write_text("body{color:#333}", encoding="utf-8")
    # 配ってはいけないもの
    (root / "server").mkdir()
    (root / "server" / ".env").write_text("API_KEY=ひみつ", encoding="utf-8")
    (root / "server" / "manage.py").write_text("# Django", encoding="utf-8")
    (root / "server" / "backups").mkdir()
    (root / "server" / "backups" / "会員.csv").write_text("MYM-0001", encoding="utf-8")
    (root / "docs").mkdir()
    (root / "docs" / "秘密の点検.md").write_text("# 秘密", encoding="utf-8")
    (root / "gas").mkdir()
    (root / "gas" / "Code.gs").write_text("// GAS", encoding="utf-8")
    (root / ".env").write_text("ひみつ", encoding="utf-8")
    settings.PREVIEW_APP_DIR = str(root)
    return root


def test_3つのタブが出る(as_owner, 写し):
    for url in ["/manage/preview/app/", "/manage/preview/bijiris/", "/manage/preview/reserve/"]:
        page = as_owner.get(url)
        assert page.status_code == 200, url
        html = page.content.decode()
        i = html.find('<nav class="page-tabs"')
        tabs = html[i:html.find("</nav>", i)]
        for t in ["まゆみ助産院アプリ", "ビジリス", "予約システム"]:
            assert t in tabs, (url, t)
    # 開いているタブが当たっている
    assert 'class="active">まゆみ助産院アプリ</a>' in as_owner.get("/manage/preview/app/").content.decode()
    assert 'class="active">予約システム</a>' in as_owner.get("/manage/preview/reserve/").content.decode()
    # ビジリス管理のタブからは外れた（引っ越したので二重に出さない）
    html = as_owner.get("/manage/bijiris/customers/").content.decode()
    i = html.find('<nav class="page-tabs"')
    assert "お客様の画面" not in html[i:html.find("</nav>", i)]


def test_切り替えが効く(as_owner, 写し):
    def 枠の中(url):
        """枠（iframe）に何を出しているか。左メニューにも似た住所が出るので、src だけを見る。"""
        html = as_owner.get(url).content.decode()
        i = html.find('<iframe class="hp-pv"')
        return html[i:html.find(">", i)]

    # 既定は「いま作っている方」。枠の中は手元のファイルの配り口
    assert 'src="/manage/preview/files/app/"' in 枠の中("/manage/preview/app/")
    assert 'src="/manage/preview/files/bijiris/"' in 枠の中("/manage/preview/bijiris/")
    # 「お客様に出ている方」は公開先の住所がそのまま埋まっている
    assert f'src="{アプリ公開先}"' in 枠の中("/manage/preview/app/?which=live")
    assert f'src="{ビジリス公開先}"' in 枠の中("/manage/preview/bijiris/?which=live")
    # 幅は 375px（スマホ）とパソコン（枠いっぱい）
    assert "width:375px" in as_owner.get("/manage/preview/app/?w=375").content.decode()
    assert "width:100%" in as_owner.get("/manage/preview/app/?w=0").content.decode()
    # 知らない値は既定に戻す（落ちない）
    html = as_owner.get("/manage/preview/app/?which=よそ&w=9999").content.decode()
    assert "width:375px" in html and "/manage/preview/files/app/" in html
    # 「別の窓で開く」がどちらにも付いている
    assert "別の窓で開く" in html and "別の窓で開く" in as_owner.get("/manage/preview/app/?which=live").content.decode()


def test_予約システムは住所をそのまま出す(as_owner, settings, 写し):
    settings.PREVIEW_RESERVE_DEV_URL = "https://例.test:10004/"
    settings.PREVIEW_RESERVE_LIVE_URL = "https://例.test:10000/"
    html = as_owner.get("/manage/preview/reserve/").content.decode()
    # ファイルは配らない。確認用の住所を枠に出す
    assert 'src="https://例.test:10004/"' in html and "/manage/preview/files/" not in html
    # 枠に出ないことがあるので、そのことと「別の窓で開く」を必ず出す
    assert "別の窓で開く" in html and "出ないことがあります" in html
    assert "https://例.test:10000/" in as_owner.get("/manage/preview/reserve/?which=live").content.decode()
    # 予約システムはファイルを配る口を持たない
    assert as_owner.get("/manage/preview/files/reserve/").status_code == 404


def test_手元のファイルを配り枠に出せる(as_owner, 写し):
    r = as_owner.get("/manage/preview/files/app/")
    assert r.status_code == 200 and "お客様アプリ".encode() in b"".join(r.streaming_content)
    # 枠（iframe）に出せること。既定の DENY だと「接続が拒否されました」になる
    assert r.get("X-Frame-Options", "") != "DENY"
    assert as_owner.get("/manage/preview/files/app/style.css")["Content-Type"] == "text/css"
    assert as_owner.get("/manage/preview/files/app/icons/icon-192.png").status_code == 200
    r = as_owner.get("/manage/preview/files/bijiris/")
    assert r.status_code == 200 and "アンケート".encode() in b"".join(r.streaming_content)
    assert r.get("X-Frame-Options", "") != "DENY"
    assert as_owner.get("/manage/preview/files/bijiris/styles.css")["Content-Type"] == "text/css"
    assert as_owner.get("/manage/preview/files/app/nothing.js").status_code == 404
    assert as_owner.get("/manage/preview/files/よそ/index.html").status_code == 404


def test_上の道は弾く(as_owner, 写し):
    """置き場の外は出さない。"""
    for kind in ["app", "bijiris"]:
        for path in ["../.env", "icons/../../.env", "..%2F.env", "a/../../.env"]:
            assert as_owner.get(f"/manage/preview/files/{kind}/" + path).status_code == 404, (kind, path)


def test_サーバーの中身は配らない(as_owner, 写し):
    """**この Django のコード・.env・backups・docs・gas は絶対に出さない。**通してよい一覧で止まる。"""
    for path in ["server/.env", "server/manage.py", "server/backups/会員.csv", ".env",
                 "docs/秘密の点検.md", "gas/Code.gs", "server/", "server/backups/"]:
        assert as_owner.get("/manage/preview/files/app/" + path).status_code == 404, path
    # ビジリスは customer-app の下だけ。その外（写しの根っこ）へは出られない
    assert as_owner.get("/manage/preview/files/bijiris/index.html").status_code == 200
    assert as_owner.get("/manage/preview/files/bijiris/server/.env").status_code == 404


def test_写しが無くても画面が開く(as_owner, settings, tmp_path):
    settings.PREVIEW_APP_DIR = str(tmp_path / "ない")
    for url, 公開先 in [("/manage/preview/app/", アプリ公開先), ("/manage/preview/bijiris/", ビジリス公開先)]:
        page = as_owner.get(url)
        assert page.status_code == 200, url
        html = page.content.decode()
        # 落ちずに、手元が無いことを伝え、公開の方は見られる
        assert "手元にファイルがありません" in html and 公開先 in html
    assert as_owner.get("/manage/preview/files/app/").status_code == 404
    # 予約システムはファイルを見ないので、写しが無くても今までどおり
    assert "手元にファイルがありません" not in as_owner.get("/manage/preview/reserve/").content.decode()
    # 設定そのものが空でも落ちない
    settings.PREVIEW_APP_DIR = ""
    assert as_owner.get("/manage/preview/app/").status_code == 200


def test_古い住所から転送する(as_owner, 写し):
    """ビジリスの古い住所（/manage/bijiris/preview/）は新しい住所へ送る。"""
    r = as_owner.get("/manage/bijiris/preview/")
    assert r.status_code == 302 and r["Location"] == "/manage/preview/bijiris/"
    r = as_owner.get("/manage/bijiris/preview/files/styles.css")
    assert r.status_code == 302 and r["Location"] == "/manage/preview/files/bijiris/styles.css"


def test_スタッフは403で未ログインはログインへ(client, staff):
    urls = ["/manage/preview/app/", "/manage/preview/bijiris/", "/manage/preview/reserve/",
            "/manage/preview/files/app/", "/manage/preview/files/bijiris/"]
    for url in urls:
        r = client.get(url)
        assert r.status_code == 302 and r["Location"].startswith("/manage/login/"), url
    client.force_login(staff)
    for url in urls:
        assert client.get(url).status_code == 403, url
