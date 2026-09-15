"""公式サイト「✍️ まゆみのつぶやき」（apps/hp/views_blog.py）。中身は mayumi-site/admin/app.py の sec_blog と同じ。"""

import datetime
import io
import json

import pytest

pytestmark = pytest.mark.django_db


def _blog(site_repo):
    return json.loads((site_repo / "admin" / "blog.json").read_text(encoding="utf-8"))


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (1400, 700), (200, 120, 80)).save(buf, "PNG")
    return buf.getvalue()


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get("/manage/hp/blog/")
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get("/manage/hp/blog/").status_code in (302, 403)


def test_一覧_件数とページ送りと検索(as_owner, site_repo):
    posts = _blog(site_repo)["posts"]
    page = as_owner.get("/manage/hp/blog/").content.decode()
    assert "まゆみのつぶやき" in page and "全 <b>%d</b> 記事" % len(posts) in page
    assert "＋ 新しく書く" in page and "ブログを見る" in page and "記事の一覧" in page
    # 旧アプリと同じ 20 件ずつ・新しい順
    新しい順 = sorted(posts, key=lambda p: (p["date"], p["slug"]), reverse=True)
    assert "/manage/hp/blog/%s/" % 新しい順[0]["slug"] in page
    assert "/manage/hp/blog/%s/" % 新しい順[20]["slug"] not in page
    pages = (len(posts) + 19) // 20
    assert "page=%d" % pages in page and "…" in page
    r = as_owner.get("/manage/hp/blog/?page=2")
    assert "/manage/hp/blog/%s/" % 新しい順[20]["slug"] in r.content.decode()
    # 検索は タイトル・本文 の両方。見つからなければその旨
    word = 新しい順[5]["title"][:4]
    hit = [p for p in posts if word.lower() in p["title"].lower() or word.lower() in p.get("text", "").lower()]
    page = as_owner.get("/manage/hp/blog/", {"q": word}).content.decode()
    assert "全 <b>%d</b> 記事" % len(hit) in page and "/manage/hp/blog/%s/" % 新しい順[5]["slug"] in page
    page = as_owner.get("/manage/hp/blog/", {"q": "zzz_どこにもない語_zzz"}).content.decode()
    assert "見つかりませんでした。" in page and "全 <b>0</b> 記事" in page


def test_新しく書く_blogjsonの形は旧アプリと同じ(as_owner, site_repo):
    before = _blog(site_repo)["posts"]
    r = as_owner.post("/manage/hp/blog/", {"action": "new"}, follow=True)
    page = r.content.decode()
    assert "新しい記事を作りました" in page and "新しいつぶやき" in page
    today = datetime.date.today()
    d = _blog(site_repo)
    p = d["posts"][0]  # 先頭に入る
    assert len(d["posts"]) == len(before) + 1
    assert list(p.keys()) == ["slug", "date", "date_disp", "title", "text", "images", "src"]
    assert p["slug"].startswith(today.strftime("%Y%m%d") + "-") and p["date"] == today.isoformat()
    assert p["date_disp"] == "%d.%d.%d" % (today.year, today.month, today.day)
    assert p["title"] == "（新しいつぶやき）" and p["images"] == [] and p["src"] == ""
    # 書き方（indent=1・日本語そのまま）も blog.save のまま
    raw = (site_repo / "admin" / "blog.json").read_text(encoding="utf-8")
    assert raw.startswith('{\n "posts": [\n  {\n   "slug": "') and "（新しいつぶやき）" in raw
    # 記事ページも作られている
    assert (site_repo / "blog" / p["slug"] / "index.html").is_file()
    # 同じ日にもう1本 → 連番
    as_owner.post("/manage/hp/blog/", {"action": "new"})
    slugs = [x["slug"] for x in _blog(site_repo)["posts"][:2]]
    assert slugs[0] != slugs[1] and all(s.startswith(today.strftime("%Y%m%d") + "-") for s in slugs)


def test_編集画面の項目(as_owner, site_repo):
    p = _blog(site_repo)["posts"][1]
    page = as_owner.get("/manage/hp/blog/%s/" % p["slug"]).content.decode()
    for s in ("記事の編集：" + p["title"][:30], "日付（表示）", "日付（機械用・2026-04-13 の形）", "タイトル",
              "URLの文字（英数字。変えると前のURLは開けなくなります）", "本文", "太字", "リンク", "選んだ文字を囲みます",
              "1行が1段落になります。", "{img:ファイル名}", "この記事の写真", "自動でWebP化・縮小",
              "保存してページを作り直す", "プレビューで確認", "この記事を削除", "閉じる"):
        assert s in page, s
    assert 'value="%s"' % p["date"] in page and 'value="%s"' % p["slug"] in page
    assert as_owner.get("/manage/hp/blog/nai-slug/").status_code == 404


def test_保存_URLの文字の変更と重なりの拒否(as_owner, site_repo):
    posts = _blog(site_repo)["posts"]
    p, other = posts[1], posts[2]
    入力 = {"action": "save", "date_disp": "2026.1.2", "date": "2026-01-02", "title": "直したタイトル",
          "text": "1行目\r\n\r\n2行目", "images": "", "new_slug": "test-slug!!"}
    r = as_owner.post("/manage/hp/blog/%s/" % p["slug"], 入力, follow=True)
    page = r.content.decode()
    assert "保存しました。" in page and "ページを作り直しました（記事%d本／一覧" % len(posts) in page
    assert r.redirect_chain[-1][0].endswith("/manage/hp/blog/test-slug/")  # 英数字以外は落ちる
    d = _blog(site_repo)
    saved = [x for x in d["posts"] if x["slug"] == "test-slug"][0]
    assert saved["title"] == "直したタイトル" and saved["text"] == "1行目\n\n2行目" and saved["date"] == "2026-01-02"
    assert saved["src"] == p["src"] and list(saved.keys()) == list(p.keys())  # 画面に無い項目と並びは残る
    assert d["posts"][-1]["slug"] == "test-slug" and not any(x["slug"] == p["slug"] for x in d["posts"])  # 旧アプリと同じく末尾へ
    assert (site_repo / "blog" / "test-slug" / "index.html").is_file()
    assert not (site_repo / "blog" / p["slug"]).exists()
    # 他の記事と同じ URL にはできない
    入力["new_slug"] = other["slug"]
    r = as_owner.post("/manage/hp/blog/test-slug/", 入力)
    assert r.status_code == 200 and "同じURLの記事がすでにあります" in r.content.decode()
    assert [x for x in _blog(site_repo)["posts"] if x["slug"] == "test-slug"]


def test_写真の追加は保存前にファイルだけ置く(as_owner, site_repo):
    p = _blog(site_repo)["posts"][1]
    from django.core.files.uploadedfile import SimpleUploadedFile

    r = as_owner.post("/manage/hp/blog/%s/" % p["slug"], {
        "action": "addimg", "date_disp": p["date_disp"], "date": p["date"], "title": "写真つき",
        "text": p["text"], "images": "\n".join(p["images"]),
        "file": SimpleUploadedFile("Photo 1.png", _png(), content_type="image/png")})
    page = r.content.decode()
    assert "写真を追加しました" in page and "未保存の変更があります" in page
    assert 'value="写真つき"' in page  # 入力中の内容はそのまま
    img = site_repo / "assets" / "img" / "blog" / "blog_Photo1.webp"
    assert img.is_file()
    from PIL import Image

    assert Image.open(img).width == 1200  # 旧アプリと同じ幅に縮める
    assert "blog_Photo1.webp" in page
    # まだ blog.json には付いていない
    assert "blog_Photo1.webp" not in _blog(site_repo)["posts"][1]["images"]
    # 保存で付く
    as_owner.post("/manage/hp/blog/%s/" % p["slug"], {
        "action": "save", "date_disp": p["date_disp"], "date": p["date"], "title": "写真つき",
        "text": p["text"], "images": "\n".join(p["images"] + ["blog_Photo1.webp"]), "new_slug": p["slug"]})
    saved = [x for x in _blog(site_repo)["posts"] if x["slug"] == p["slug"]][0]
    assert saved["images"][-1] == "blog_Photo1.webp"
    assert "blog_Photo1.webp" in (site_repo / "blog" / p["slug"] / "index.html").read_text(encoding="utf-8")
    # 写真は手元のサイトから見られる
    assert as_owner.get("/manage/hp/site/assets/img/blog/blog_Photo1.webp").status_code == 200


def test_プレビューは保存せずに見られる(as_owner, site_repo):
    p = _blog(site_repo)["posts"][1]
    r = as_owner.post("/manage/hp/blog/%s/" % p["slug"], {
        "action": "preview", "date_disp": p["date_disp"], "date": p["date"], "title": "プレビューの題",
        "text": "プレビューの本文", "images": "", "new_slug": p["slug"]})
    page = r.content.decode()
    assert "いま入力している内容で表示しています（保存はされていません）" in page
    assert "/manage/hp/preview/files/blog/%s/" % p["slug"] in page and "入力を反映して更新" in page
    html = (site_repo / "admin" / "_preview" / "blog" / p["slug"] / "index.html").read_text(encoding="utf-8")
    assert "プレビューの題" in html and "プレビューの本文" in html
    assert [x for x in _blog(site_repo)["posts"] if x["slug"] == p["slug"]][0]["title"] == p["title"]
    r = as_owner.get("/manage/hp/preview/files/blog/%s/" % p["slug"])
    assert r.status_code == 200 and "プレビューの題".encode() in b"".join(r.streaming_content)


def test_削除(as_owner, site_repo):
    posts = _blog(site_repo)["posts"]
    p = posts[1]
    r = as_owner.post("/manage/hp/blog/%s/" % p["slug"], {"action": "del"}, follow=True)
    page = r.content.decode()
    assert "削除しました（%d → %d 記事）。" % (len(posts), len(posts) - 1) in page and "ページを作り直しました。" in page
    assert not any(x["slug"] == p["slug"] for x in _blog(site_repo)["posts"])
    assert not (site_repo / "blog" / p["slug"]).exists()
