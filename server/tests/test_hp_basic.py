"""公式サイト（apps/hp）: 🏥 基本情報。中身は mayumi-site/admin/app.py の sec_basic_page と同じ。"""

import json

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

pytestmark = pytest.mark.django_db

URL = "/manage/hp/basic/"
SAVE = "/manage/hp/basic/save/"


def _read(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get(URL)
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()
    r = as_owner.post(SAVE, {"part": "basic"})
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(URL).status_code in (302, 403)
    assert client.post(SAVE, {"part": "basic"}).status_code in (302, 403)


def test_画面に旧アプリと同じ項目が並ぶ(as_owner, site_repo):
    d = _read(site_repo)
    page = as_owner.get(URL).content.decode()
    # 目次と各部の見出し（旧アプリの BASIC_PARTS の順）
    for s in ("このページの中身", "🏥 基本情報", "🕒 診療時間", "💰 料金表", "🔖 SNS・ブログ", "📄 産後ケアの資料"):
        assert s in page, s
    assert page.index("sec-hours") < page.index("sec-prices") < page.index("sec-sns") < page.index("sec-doc")
    # 基本情報: 11 項目のラベルと、いまの値
    for label in ("電話番号", "LINE予約のURL", "住所（正式表記）", "フッターの案内文", "ページ上部のオレンジ帯の文"):
        assert label in page, label
    assert d["values"]["tel"] in page and "電話番号は51か所" in page
    # 診療時間: 表と注記と1行表記
    assert d["hours"]["rows"][0][0] in page and "休診日などの注記" in page and "📝 1行表記" in page
    # 料金表: 表ごとの見出し、まとめた表は読むだけ
    assert "母乳外来：乳房マッサージ" in page and "母乳外来：まとめた表（料金ページ用）" in page
    assert "中身（読むだけ）" in page and "＋ 行を追加" in page and "👁 プレビュー" in page
    assert "表の見出しもここで直せます" in page
    # SNS・Facebook・つぶやき・資料
    assert "＋ リンクを新しく追加" in page and "名前（管理用）" in page and "ロゴの幅（px）" in page
    assert "📘 Facebookの投稿をトップページに出す" in page and "👁 いまの設定で見てみる" in page
    assert "✍️ まゆみのつぶやきをトップページに出す" in page and "出す記事の数" in page
    assert "📄 資料を見るボタン（産後ケアのページ）" in page and "PDFを入れ替える" in page
    # 写しには PDF が無いので「見つかりません」
    assert "が見つかりません" in page
    # 左メニューに「基本情報」
    assert 'class="active">基本情報</a>' in page  # 上部タブ（左メニューは「ページを整える」にまとめた）


def test_基本情報の保存はvaluesだけを書く(as_owner, site_repo):
    before = _read(site_repo)
    data = {"part": "basic"}
    for k, v in before["values"].items():
        data["v-" + k] = v
    data["v-tel"] = "070-0000-0000"
    data["v-tagline"] = "新しい帯の文"
    r = as_owner.post(SAVE, data, follow=True)
    page = r.content.decode()
    assert "下書きを保存しました（サイトにはまだ反映していません）。" in page
    assert 'value="070-0000-0000"' in page and "新しい帯の文" in page
    after = _read(site_repo)
    assert after["values"]["tel"] == "070-0000-0000" and after["values"]["tagline"] == "新しい帯の文"
    # ほかの部分には触らない
    for k in ("hours", "prices", "sns", "sns_embed", "blog_top", "page_heads", "docs", "news_top_count"):
        assert after.get(k) == before.get(k), k


def test_診療時間の保存(as_owner, site_repo):
    before = _read(site_repo)
    r = as_owner.post(SAVE, {
        "part": "hours", "h-0-0": "平日", "h-0-1": "10時〜16時", "h-1-0": "土曜", "h-1-1": "9時〜12時",
        "hnote": "休診日：水曜\n※臨時休診あり", "v-hours_cta": "平日 10時〜16時／土 9時〜12時",
    }, follow=True)
    assert "下書きを保存しました" in r.content.decode()
    after = _read(site_repo)
    assert after["hours"]["rows"][0] == ["平日", "10時〜16時"] and after["hours"]["rows"][1] == ["土曜", "9時〜12時"]
    assert after["hours"]["note"] == "休診日：水曜\n※臨時休診あり"
    assert after["values"]["hours_cta"] == "平日 10時〜16時／土 9時〜12時"
    # values のほかの項目はそのまま
    assert after["values"]["tel"] == before["values"]["tel"]


def test_料金表の保存と行の追加と削除(as_owner, site_repo):
    key = "bonyu-massage"
    d = _read(site_repo)
    rows = d["prices"][key]["rows"]
    base = {"part": "price", "key": key, "caption": "乳房マッサージ（改）", "note": "※税別",
            "nrows": len(rows), "ncol": 2}
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            base["cell-%d-%d" % (i, j)] = c
    base["cell-0-1"] = "5,500円（税別）"
    r = as_owner.post(SAVE, base, follow=True)
    assert "下書きを保存しました" in r.content.decode()
    p = _read(site_repo)["prices"][key]
    assert p["caption"] == "乳房マッサージ（改）" and p["note"] == "※税別" and p["rows"][0][1] == "5,500円（税別）"
    # 行を追加 → 空の行が末尾に増える
    as_owner.post(SAVE, {**base, "part": "price_row_add"})
    p = _read(site_repo)["prices"][key]
    assert len(p["rows"]) == len(rows) + 1 and p["rows"][-1] == ["", ""]
    # 行を削除 → 指定の行が消える
    as_owner.post(SAVE, {**base, "part": "price_row_del", "row": 0})
    p = _read(site_repo)["prices"][key]
    assert len(p["rows"]) == len(rows) - 1 and p["rows"][0][0] == rows[1][0]
    # 無い表は断る
    r = as_owner.post(SAVE, {"part": "price", "key": "nai"}, follow=True)
    assert "その料金表はありません" in r.content.decode()


def test_料金表のまとめた表は見出しだけ変わる(as_owner, site_repo):
    before = _read(site_repo)
    r = as_owner.post(SAVE, {"part": "price", "key": "bonyu-all", "caption": "母乳外来（全部）", "note": "",
                             "nrows": 1, "ncol": 2, "cell-0-0": "勝手な行", "cell-0-1": "0円"}, follow=True)
    assert r.status_code == 200
    p = _read(site_repo)["prices"]["bonyu-all"]
    assert p["caption"] == "母乳外来（全部）" and p["rows"] == before["prices"]["bonyu-all"]["rows"]


def test_料金表のプレビューは保存せずに表だけ返す(as_owner, site_repo):
    before = _read(site_repo)
    r = as_owner.post("/manage/hp/basic/price-preview/", {
        "key": "itothermy", "caption": "見出しの試し", "note": "※注記", "nrows": 1, "ncol": 2,
        "cell-0-0": "試しの項目", "cell-0-1": "1,234円",
    })
    page = r.content.decode()
    assert r.status_code == 200 and "<caption>見出しの試し" in page and "※注記" in page
    assert "<th>試しの項目</th>" in page and "1,234円" in page and "まだ保存していません" in page
    assert _read(site_repo) == before
    # まとめた表は元の表から行を作る
    r = as_owner.post("/manage/hp/basic/price-preview/", {"key": "bonyu-all", "caption": "母乳外来", "note": ""})
    assert "乳房マッサージ" in r.content.decode()
    assert as_owner.post("/manage/hp/basic/price-preview/", {"key": "nai"}).status_code == 404


def test_SNSの追加と保存と並べ替えと削除(as_owner, site_repo):
    n = len(_read(site_repo)["sns"])
    r = as_owner.post(SAVE, {"part": "sns_add"}, follow=True)
    assert "下書きを保存しました" in r.content.decode()
    sns = _read(site_repo)["sns"]
    assert len(sns) == n + 1 and sns[-1]["name"] == "新しいリンク" and sns[-1]["url"] == "https://"
    assert sns[-1]["btn_class"] == "sns__other" and sns[-1]["show"] is True
    # 保存: 名前・URL・幅・ボタンの文字・表示
    r = as_owner.post(SAVE, {"part": "sns", "i": n, "s-name": "X", "s-url": "https://x.com/mayumi",
                             "s-icon_w": "abc", "s-btn_label": "X（旧Twitter）", "s-show": "0"}, follow=True)
    s = _read(site_repo)["sns"][n]
    assert s["name"] == "X" and s["url"] == "https://x.com/mayumi" and s["icon_w"] == 44
    assert s["btn_label"] == "X（旧Twitter）" and s["show"] is False
    # 上へ動かす（画面の入力も一緒に保存される）
    as_owner.post(SAVE, {"part": "sns_move", "i": n, "dir": "-1", "s-name": "X2", "s-url": "https://x.com/mayumi",
                         "s-icon_w": "50", "s-btn_label": "X", "s-show": "1"})
    sns = _read(site_repo)["sns"]
    assert sns[n - 1]["name"] == "X2" and sns[n - 1]["icon_w"] == 50 and sns[n - 1]["show"] is True
    # 端を越えては動かない
    as_owner.post(SAVE, {"part": "sns_move", "i": 0, "dir": "-1"})
    assert len(_read(site_repo)["sns"]) == n + 1
    # 削除
    as_owner.post(SAVE, {"part": "sns_del", "i": n - 1})
    sns = _read(site_repo)["sns"]
    assert len(sns) == n and all(s["name"] != "X2" for s in sns)
    # 無い番号は断る
    r = as_owner.post(SAVE, {"part": "sns_del", "i": 99}, follow=True)
    assert "見つかりませんでした" in r.content.decode() and len(_read(site_repo)["sns"]) == n


def test_SNSのロゴの差し替え(as_owner, site_repo):
    from PIL import Image
    import io

    buf = io.BytesIO()
    Image.new("RGB", (80, 40), (200, 100, 50)).save(buf, "PNG")
    r = as_owner.post(SAVE, {"part": "sns_icon", "target": "sns-test.png",
                             "file": SimpleUploadedFile("logo.png", buf.getvalue(), content_type="image/png")},
                      follow=True)
    assert "sns-test.png を差し替えました" in r.content.decode()
    assert (site_repo / "assets" / "img" / "sns-test.png").is_file()
    # ファイル名が決まっていないリンクには入れられない
    r = as_owner.post(SAVE, {"part": "sns_icon", "target": "",
                             "file": SimpleUploadedFile("logo.png", buf.getvalue(), content_type="image/png")},
                      follow=True)
    assert "ロゴのファイル名がありません" in r.content.decode()


def test_Facebookの欄とつぶやきの欄の保存(as_owner, site_repo):
    r = as_owner.post(SAVE, {"part": "sns_embed", "fb-show": "0", "fb-url": "https://www.facebook.com/x",
                             "fb-height": "", "fb-title": "説明"}, follow=True)
    assert "下書きを保存しました" in r.content.decode()
    c = _read(site_repo)["sns_embed"]
    assert c == {"show": False, "url": "https://www.facebook.com/x", "height": 560, "title": "説明"}
    r = as_owner.post(SAVE, {"part": "blog_top", "bt-show": "1", "bt-count": "4", "bt-height": "600"}, follow=True)
    c = _read(site_repo)["blog_top"]
    assert c == {"show": True, "count": 4, "height": 600}
    # 空なら旧アプリと同じ既定値（6 と 560）
    as_owner.post(SAVE, {"part": "blog_top", "bt-show": "1", "bt-count": "", "bt-height": "x"})
    assert _read(site_repo)["blog_top"] == {"show": True, "count": 6, "height": 560}


def test_資料の設定の保存とPDFの入れ替え(as_owner, site_repo):
    r = as_owner.post(SAVE, {"part": "doc", "doc-show": "0", "doc-label": "資料を読む", "doc-note": "PDF"}, follow=True)
    assert "下書きを保存しました" in r.content.decode()
    c = _read(site_repo)["docs"]["bijiris"]
    assert c["show"] is False and c["label"] == "資料を読む" and c["note"] == "PDF"
    assert c["file"] == "bijiris-text.pdf"  # ファイル名は触らない
    # PDF を入れる → assets/doc/bijiris-text.pdf に決まった名前で置かれる
    pdf = b"%PDF-1.4\n%test\n"
    r = as_owner.post(SAVE, {"part": "doc_upload", "key": "bijiris",
                             "file": SimpleUploadedFile("ビジリス資料.pdf", pdf, content_type="application/pdf")},
                      follow=True)
    page = r.content.decode()
    assert "PDF を入れ替えました" in page
    assert (site_repo / "assets" / "doc" / "bijiris-text.pdf").read_bytes() == pdf
    assert _read(site_repo)["docs"]["bijiris"]["file"] == "bijiris-text.pdf"
    assert "いま置いてあるファイル：<b>bijiris-text.pdf</b>" in page and "いまのPDFを開いてみる" in page
    # 変な key は文字を落として決まった名前にする（別の場所には書かない）
    as_owner.post(SAVE, {"part": "doc_upload", "key": "../x",
                         "file": SimpleUploadedFile("a.pdf", pdf, content_type="application/pdf")})
    assert (site_repo / "assets" / "doc" / "x-text.pdf").is_file()
    # ファイル無し
    r = as_owner.post(SAVE, {"part": "doc_upload", "key": "bijiris"}, follow=True)
    assert "ファイルが受け取れませんでした" in r.content.decode()


def test_知らない部分は断る(as_owner, site_repo):
    before = _read(site_repo)
    r = as_owner.post(SAVE, {"part": "nandemo"}, follow=True)
    assert "何を保存するのか分かりませんでした" in r.content.decode()
    assert _read(site_repo) == before
