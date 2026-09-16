"""公式サイト（apps/hp）: 📢 お知らせ。中身は mayumi-site/admin/app.py の sec_news / news-* と同じ。"""

import datetime
import io
import json

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.content.models import News

pytestmark = pytest.mark.django_db


def _読む(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (1600, 800), (200, 120, 80)).save(buf, format="PNG")
    return buf.getvalue()


def _保存の入力(item, **extra):
    """編集欄の入力（nGrab が集めるのと同じ項目）。"""
    v = {"action": "save", "date": item["date"], "title": item["title"],
         "category": item.get("category", ""), "text": item.get("text", ""),
         "images": item.get("images", [])}
    v.update(extra)
    return v


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    for u in ("/manage/hp/news/", "/manage/hp/news/0/", "/manage/hp/news/app/"):
        r = as_owner.get(u)
        assert r.status_code == 200 and "まだ設定されていません" in r.content.decode(), u


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get("/manage/hp/news/").status_code in (302, 403)


def test_一覧_項目と文言(as_owner, site_repo):
    d = _読む(site_repo)
    page = as_owner.get("/manage/hp/news/").content.decode()
    for s in ("📢 お知らせ", "トップに出す件数", "＋ 新しく書く", "お知らせを見る", "📥 アプリのNEWSから取り込む",
              "お知らせの一覧", "↑↓で並べ替えられます", "種別でしぼる", "（種別なし）",
              "全 <b>%d</b> 件" % len(d["news"])):
        assert s in page, s
    for n in d["news"]:
        assert n["title"] in page and n["date_disp"] in page
    # 種別でしぼる（pickBy と同じ: 一致・種別なし）
    page = as_owner.get("/manage/hp/news/?cat=イベント情報").content.decode()
    有 = [n for n in d["news"] if n.get("category") == "イベント情報"]
    無 = [n for n in d["news"] if n.get("category") != "イベント情報"]
    assert all(n["title"] in page for n in 有) and all(n["title"] not in page for n in 無)
    assert "%d件 / 全%d件" % (len(有), len(d["news"])) in page
    page = as_owner.get("/manage/hp/news/?cat=ないもの").content.decode()
    assert "その種別のお知らせはありません。" in page
    # 左メニューの「公式サイト」の段に出る
    assert 'href="/manage/hp/news/"' in page and 'class="active">お知らせ（サイト）</a>' in page  # 上部タブ


def test_トップに出す件数を保存(as_owner, site_repo):
    r = as_owner.post("/manage/hp/news/", {"action": "topcount", "news_top_count": "5"}, follow=True)
    assert "トップに出す件数を保存しました。" in r.content.decode()
    assert _読む(site_repo)["news_top_count"] == 5
    # 数でなければ 3 に戻す（旧アプリの parseInt(...)||3 と同じ）
    as_owner.post("/manage/hp/news/", {"action": "topcount", "news_top_count": "abc"})
    assert _読む(site_repo)["news_top_count"] == 3


def test_新しく書く_先頭に足して編集欄が開く(as_owner, site_repo):
    before = _読む(site_repo)
    r = as_owner.post("/manage/hp/news/", {"action": "new"})
    assert r.status_code == 302 and r.url == "/manage/hp/news/0/"
    d = _読む(site_repo)
    n = d["news"][0]
    assert len(d["news"]) == len(before["news"]) + 1
    assert n["title"] == "（新しいお知らせ）" and n["date"] == datetime.date.today().isoformat()
    assert n["blocks"] == [{"t": "p", "v": "ここに本文を書いてください。"}] and n["images"] == []
    assert n["id"] not in {x["id"] for x in before["news"]}
    page = as_owner.get("/manage/hp/news/0/", follow=True).content.decode()
    assert "新しいお知らせを作りました。書けたら「保存してページを作り直す」を押してください。" in page
    for s in ("新しいお知らせ", "日付", "種別（タイトルの前に出ます）", "タイトル", "写真", "写真はまだありません。",
              "本文に差し込む", "本文", "見出しにする", "小見出しにする", "箇条書きにする", "本文にもどす",
              "保存してページを作り直す", "プレビューで確認", "このお知らせを削除", "閉じる",
              "種別は「選択肢の設定」で増やせます。", "改行するとサイトでも改行されます。"):
        assert s in page, s
    # 昔の形（blocks）のままでも、本文が編集欄に出る
    assert "ここに本文を書いてください。</textarea>" in page
    for c in ("イベント情報", "休診情報", "商品情報"):
        assert '<option value="%s"' % c in page


def test_昔の形のお知らせも本文が編集欄に出る(as_owner, site_repo):
    d = _読む(site_repo)
    i, n = next((i, n) for i, n in enumerate(d["news"]) if not n.get("text") and n.get("blocks"))
    page = as_owner.get("/manage/hp/news/%d/" % i).content.decode()
    first = n["blocks"][0]["v"].split("\n")[0]
    assert first in page and "お知らせの編集：" + n["title"][:30] in page


def test_保存してページを作り直す(as_owner, site_repo):
    as_owner.post("/manage/hp/news/", {"action": "new"})
    d = _読む(site_repo)
    n = d["news"][0]
    r = as_owner.post("/manage/hp/news/0/", _保存の入力(
        n, date="2026-10-03", title="味噌作り教室", category="イベント情報",
        text="# 見出し\r\n\r\n本文です。\r\n二行目\r\n\r\n- ひとつ\r\n- ふたつ"), follow=True)
    page = r.content.decode()
    assert "保存しました。" in page and "ページを作り直しました（%d件）。" % len(d["news"]) in page
    m = _読む(site_repo)["news"][0]
    assert m["id"] == n["id"] and m["title"] == "味噌作り教室" and m["category"] == "イベント情報"
    assert m["date"] == "2026-10-03" and m["date_disp"] == "2026.10.3"
    assert "blocks" not in m and m["text"].startswith("# 見出し\n\n本文です。\n二行目")
    # 記事ページと一覧が作り直されている
    article = (site_repo / "news" / n["id"] / "index.html").read_text(encoding="utf-8")
    assert "味噌作り教室" in article and "<h3>見出し</h3>" in article and '<ul class="check-list">' in article
    assert "味噌作り教室" in (site_repo / "news" / "index.html").read_text(encoding="utf-8")
    # 編集欄の見出しも新しい題名になる
    assert "お知らせの編集：味噌作り教室" in page
    # 無い位置は一覧へ戻す
    r = as_owner.get("/manage/hp/news/99/", follow=True)
    assert "その位置のお知らせがありません。" in r.content.decode()


def test_写真を追加_保存はまだしない(as_owner, site_repo):
    as_owner.post("/manage/hp/news/", {"action": "new"})
    n = _読む(site_repo)["news"][0]
    v = _保存の入力(n, action="add_image", title="写真つき", text="本文")
    v["file"] = SimpleUploadedFile("misoshiru.png", _png(), content_type="image/png")
    r = as_owner.post("/manage/hp/news/0/", v)
    page = r.content.decode()
    assert r.status_code == 200
    assert "写真を追加しました。「本文に差し込む」で位置を決められます。" in page
    assert 'name="images" value="news_misoshiru.webp"' in page and "未保存の変更があります" in page
    assert (site_repo / "assets" / "img" / "news" / "news_misoshiru.webp").is_file()
    # 入力していた内容は編集欄に残っている。お知らせ自体はまだ保存されていない
    assert 'value="写真つき"' in page and "本文</textarea>" in page
    m = _読む(site_repo)["news"][0]
    assert m["title"] == "（新しいお知らせ）" and m["images"] == []
    # そのまま保存すると写真も入る
    as_owner.post("/manage/hp/news/0/", _保存の入力(n, title="写真つき", text="本文\n\n{img:news_misoshiru.webp}",
                                                 images=["news_misoshiru.webp"]))
    m = _読む(site_repo)["news"][0]
    assert m["images"] == ["news_misoshiru.webp"]
    article = (site_repo / "news" / n["id"] / "index.html").read_text(encoding="utf-8")
    assert "assets/img/news/news_misoshiru.webp" in article
    # 写真は手元のサイトから配られる
    assert as_owner.get("/manage/hp/site/assets/img/news/news_misoshiru.webp").status_code == 200
    # ファイルが無ければその旨
    r = as_owner.post("/manage/hp/news/0/", _保存の入力(n, action="add_image"))
    assert "ファイルが受け取れませんでした。" in r.content.decode()


def test_プレビューで確認(as_owner, site_repo):
    as_owner.post("/manage/hp/news/", {"action": "new"})
    n = _読む(site_repo)["news"][0]
    r = as_owner.post("/manage/hp/news/0/", _保存の入力(n, action="preview", title="プレビューの題", text="まだ保存しない本文"))
    page = r.content.decode()
    assert r.status_code == 200
    assert "いま入力している内容で表示しています（保存はされていません）" in page
    for s in ("スマホ", "タブレット", "PC", "入力を反映して更新"):
        assert s in page, s
    url = "/manage/hp/preview/files/news/%s/" % n["id"]
    assert 'src="%s?t=' % url in page
    html = b"".join(as_owner.get(url).streaming_content).decode("utf-8")
    assert "プレビューの題" in html and "まだ保存しない本文" in html
    # 保存はされていない
    assert _読む(site_repo)["news"][0]["title"] == "（新しいお知らせ）"


def test_削除(as_owner, site_repo):
    before = _読む(site_repo)
    gone = before["news"][0]
    r = as_owner.post("/manage/hp/news/0/", {"action": "delete"}, follow=True)
    page = r.content.decode()
    assert "「%s」を削除しました。" % gone["title"] in page and "ページを作り直しました。" in page
    d = _読む(site_repo)
    assert len(d["news"]) == len(before["news"]) - 1 and gone["id"] not in {x["id"] for x in d["news"]}
    assert not (site_repo / "news" / gone["id"]).exists()


def test_並べ替え(as_owner, site_repo):
    before = _読む(site_repo)["news"]
    r = as_owner.post("/manage/hp/news/", {"action": "move", "from": "0", "to": "1", "cat": ""}, follow=True)
    assert "並べ替えました。" in r.content.decode()
    after = _読む(site_repo)["news"]
    assert after[0]["id"] == before[1]["id"] and after[1]["id"] == before[0]["id"]
    # 範囲の外は何もしない
    as_owner.post("/manage/hp/news/", {"action": "move", "from": "0", "to": "-1"})
    assert _読む(site_repo)["news"][0]["id"] == before[1]["id"]


# ------------------------------------------------------------------ アプリの NEWS から取り込む

def _アプリのNEWSを用意(site_repo):
    """サーバーの表に、取り込み候補・取り込み済み・出さないものを置く。"""
    already = _読む(site_repo)["news"][0]
    y, m, d = (int(x) for x in already["date"].split("-"))
    News.objects.create(sheet_row=1, posted_on=datetime.date(2026, 9, 10), title="  新しい   NEWS ", category="イベント情報",
                        body="一段落目\n続き\n\n二段落目", published=True, notice_listed=True,
                        image_url='["https://drive.google.com/thumbnail?id=abc","https://drive.google.com/thumbnail?id=def"]')
    News.objects.create(sheet_row=2, posted_on=datetime.date(y, m, d), title=already["title"], category="イベント情報",
                        body="もう出したもの", published=True, notice_listed=True)
    News.objects.create(sheet_row=3, posted_on=datetime.date(2026, 9, 11), title="消した記事", category="イベント情報",
                        published=True, notice_listed=True, deleted=True)
    News.objects.create(sheet_row=4, posted_on=datetime.date(2026, 9, 12), title="一覧に出さない記事", category="休診情報",
                        published=True, notice_listed=False)
    News.objects.create(sheet_row=5, posted_on=datetime.date(2026, 9, 13), title="写真なし", category="休診情報",
                        body="本文", published=True, notice_listed=True)
    return already


def test_取り込み画面_候補と取り込み済み(as_owner, site_repo):
    already = _アプリのNEWSを用意(site_repo)
    page = as_owner.get("/manage/hp/news/app/").content.decode()
    for s in ("📥 アプリのNEWSから取り込む", "入れたいものにチェック", "まだ公式サイトに出していないものだけ",
              "すべて選ぶ", "選択を外す", "すべての種別", "選んだものを取り込む", "写真もいっしょに持ってきます。少し時間がかかります。",
              "アプリのNEWS 全3件／まだ出していないもの 2件／いま表示 2件"):
        assert s in page, s
    # 題名の空白は詰める。写真の枚数が出る。消したもの・一覧に出さないものは出ない
    assert "新しい NEWS" in page and "写真2枚" in page and "2026-09-10" in page
    assert "消した記事" not in page and "一覧に出さない記事" not in page
    # 取り込み済みは、はじめは隠れている。チェックを外すと薄く出て選べない
    assert already["title"] not in page
    page = as_owner.get("/manage/hp/news/app/?only_new=0").content.decode()
    assert already["title"] in page and "取り込み済み" in page and "disabled" in page
    assert "いま表示 3件" in page
    # 種別でしぼる
    page = as_owner.get("/manage/hp/news/app/?cat=休診情報").content.decode()
    assert "写真なし" in page and "新しい NEWS" not in page
    # 全部出し終えているとき
    News.objects.filter(sheet_row__in=(1, 5)).delete()
    page = as_owner.get("/manage/hp/news/app/").content.decode()
    assert "アプリのNEWSは、すべて公式サイトに出してあります。" in page
    page = as_owner.get("/manage/hp/news/app/?cat=ないもの&only_new=0").content.decode()
    assert "見せるものがありません。種別のしぼり込みを変えてみてください。" in page


def test_取り込む_写真も持ってくる(as_owner, site_repo, monkeypatch):
    from apps.hp import views_news

    _アプリのNEWSを用意(site_repo)
    取った = []

    def 偽の写真(url):
        取った.append(url)
        return _png()

    monkeypatch.setattr(views_news, "_写真を取る", 偽の写真)
    before = _読む(site_repo)["news"]
    # 何も選ばなければその旨
    r = as_owner.post("/manage/hp/news/app/", {}, follow=True)
    assert "取り込むものが選ばれていません。" in r.content.decode()
    r = as_owner.post("/manage/hp/news/app/", {"keys": ["2026-09-10|新しい NEWS", "2026-09-13|写真なし"]}, follow=True)
    page = r.content.decode()
    assert r.redirect_chain[-1][0] == "/manage/hp/news/"
    assert "2件を取り込みました。" in page and "取り込みました：2026-09-10 新しい NEWS" in page
    assert "取り込みました：2026-09-13 写真なし" in page and "お知らせのページを作り直しました。" in page
    assert 取った == ["https://drive.google.com/thumbnail?id=abc", "https://drive.google.com/thumbnail?id=def"]
    d = _読む(site_repo)["news"]
    assert len(d) == len(before) + 2
    # 新しい順に並び直る。日付が同じでも既存のものより前には出ない
    assert d[0]["title"] == "写真なし" and d[1]["title"] == "新しい NEWS"
    n = d[1]
    assert n["date"] == "2026-09-10" and n["date_disp"] == "2026.9.10" and n["category"] == "イベント情報"
    assert n["blocks"] == [{"t": "p", "v": "一段落目\n続き"}, {"t": "p", "v": "二段落目"}]
    assert n["images"] == ["news_app_20260910.webp", "news_app_20260910_2.webp"]
    assert n["id"] not in {x["id"] for x in before}
    for f in n["images"]:
        assert (site_repo / "assets" / "img" / "news" / f).is_file()
    assert d[0]["images"] == [] and d[0]["blocks"] == [{"t": "p", "v": "本文"}]
    # 記事ページができている
    assert "新しい NEWS" in (site_repo / "news" / n["id"] / "index.html").read_text(encoding="utf-8")
    # もう一度開くと取り込み済みになっている
    page = as_owner.get("/manage/hp/news/app/").content.decode()
    assert "アプリのNEWSは、すべて公式サイトに出してあります。" in page
    # 同じものを選び直しても二重には入らない
    r = as_owner.post("/manage/hp/news/app/", {"keys": ["2026-09-10|新しい NEWS"]}, follow=True)
    assert "取り込むものがありませんでした。" in r.content.decode()
    assert len(_読む(site_repo)["news"]) == len(before) + 2


def test_取り込む_写真が取れなくても本文は入る(as_owner, site_repo, monkeypatch):
    from apps.hp import views_news

    _アプリのNEWSを用意(site_repo)

    def 落ちる(url):
        raise OSError("timed out")

    monkeypatch.setattr(views_news, "_写真を取る", 落ちる)
    r = as_owner.post("/manage/hp/news/app/", {"keys": ["2026-09-10|新しい NEWS"]}, follow=True)
    page = r.content.decode()
    assert "1件を取り込みました。" in page and "写真を取れませんでした（新しい NEWS）: timed out" in page
    n = [x for x in _読む(site_repo)["news"] if x["title"] == "新しい NEWS"][0]
    assert n["images"] == [] and n["blocks"][0]["v"] == "一段落目\n続き"
