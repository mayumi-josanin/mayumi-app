"""公式サイト（apps/hp）: 🎒 各種お教室。中身は mayumi-site/admin/app.py の sec_classes と同じ。"""

import io
import json
import os

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

pytestmark = pytest.mark.django_db

LIST = "/manage/hp/classes/"


def _read(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def _write(site_repo, d):
    (site_repo / "admin" / "content.json").write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def _form(c, **extra):
    """編集画面が送るのと同じ形のフォーム（保存前のお教室 → 入力欄）。"""
    P = {
        "id": c.get("id", ""), "date": c.get("date", ""), "category": c.get("category", ""), "kind": c.get("kind", ""),
        "date_disp": c.get("date_disp", ""), "open_date": c.get("open_date", ""), "open_disp": c.get("open_disp", ""),
        "lead": c.get("lead", ""), "show": "0" if c.get("show") is False else "1",
        "images": list(c.get("images") or []),
    }
    for i, b in enumerate(c.get("blocks") or []):
        P["blk_t_%d" % i] = b["t"]
        P["blk_v_%d" % i] = b["v"] if isinstance(b["v"], str) else "\n".join(b["v"])
    for i, r in enumerate(c.get("reports") or []):
        P["rep_id_%d" % i] = r.get("id", "")
        P["rep_kind_%d" % i] = r.get("kind", "")
        P["rep_title_%d" % i] = r.get("title", "")
        P["rep_date_%d" % i] = r.get("date", "")
        P["rep_disp_%d" % i] = r.get("date_disp", "")
        P["rep_text_%d" % i] = r.get("text", "")
        P["rep_img_%d" % i] = list(r.get("images") or [])
    P.update(extra)
    return P


def _png(w=1400, h=700):
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (w, h), (200, 120, 80)).save(buf, "PNG")
    return SimpleUploadedFile("photo one.png", buf.getvalue(), content_type="image/png")


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    for u in (LIST, LIST + "0/"):
        r = as_owner.get(u)
        assert r.status_code == 200 and "まだ設定されていません" in r.content.decode(), u


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(LIST).status_code in (302, 403)


def test_一覧_件数と文言としぼり込み(as_owner, site_repo):
    d = _read(site_repo)
    page = as_owner.get(LIST).content.decode()
    assert "各種お教室" in page and "全 <b>%d</b> 件" % len(d["classroom"]) in page
    assert "＋ 新しいお教室" in page and "お教室のページを見る" in page and "＋ 新しく作成" in page
    assert "この順番でページに並びます。↑↓で並べ替えられます。" in page
    assert "お教室でしぼる" in page and "分類でしぼる" in page and "（名称なし）" in page and "（分類なし）" in page
    assert "全%d件" % len(d["classroom"]) in page
    for c in d["classroom"]:
        assert c["category"] in page
    assert "開催%d回" % len(d["classroom"][0]["reports"]) in page
    # 左メニューに出る
    assert "🎒</span> 各種お教室" in page
    # しぼる: 名称
    name = d["classroom"][1]["category"]
    hits = [c for c in d["classroom"] if c["category"] == name]
    page = as_owner.get(LIST, {"cat": name}).content.decode()
    assert "%d件 / 全%d件" % (len(hits), len(d["classroom"])) in page
    # 分類で該当なし
    page = as_owner.get(LIST, {"kind": "__none__"}).content.decode()
    assert "そのお教室はありません。" in page
    # 1件も無いとき
    d["classroom"] = []
    _write(site_repo, d)
    page = as_owner.get(LIST).content.decode()
    assert "まだありません。「＋ 新しいお教室」から作れます。" in page


def test_新しいお教室_先頭に足されて編集画面へ(as_owner, site_repo):
    before = _read(site_repo)["classroom"]
    r = as_owner.post(LIST, {"action": "new"})
    assert r.status_code == 302 and r["Location"] == LIST + "0/?new=1"
    cs = _read(site_repo)["classroom"]
    assert len(cs) == len(before) + 1
    # 空いている番号（class-01〜11 と class-13 があるので 12）
    assert cs[0]["id"] == "class-12" and cs[0]["blocks"] == [{"t": "p", "v": "どんなお教室かを書いてください。"}]
    page = as_owner.get(LIST + "0/?new=1").content.decode()
    assert "新しく作りました。書けたら「保存してページを作り直す」を押してください。" in page
    assert "新しいお教室" in page


def test_編集画面_項目と文言(as_owner, site_repo):
    c = _read(site_repo)["classroom"][0]
    page = as_owner.get(LIST + "0/").content.decode()
    assert "お教室の編集：" + c["category"] in page
    for s in ("お教室の名称", "分類", "開催のしかた（空でも可）", "開催日（空でも可）", "表示する",
              "一覧に出す短い説明（空でも可）", "教室の写真", "1枚目が一覧カードのサムネイルになります。",
              "教室の概要", "＋段落", "＋小見出し", "＋チェックリスト", "＋囲み枠",
              "リンクは <b>[表示する文字](URL)</b>、太字は <b>**はさむ**</b> と書きます。",
              "開催の予定・これまでの開催", "＋ 開催予定を作成", "＋ 開催した回を足す", "新しい順に並べ替える",
              "足した回はいちばん上に入ります。", "保存してページを作り直す", "プレビューで確認", "これを削除", "閉じる",
              "写真はまだありません。", "1行あけると段落が分かれます。", "表示する日付"):
        assert s in page, s
    # 回の番号は上が大きい（旧アプリと同じ）。分類は日付から決まる
    n = len(c["reports"])
    assert "%d回目：%s" % (n, c["reports"][0]["title"]) in page
    assert "開催予定" in page and "過去の開催" in page and "開催日を過ぎたため" in page
    # 無い位置は 404
    assert as_owner.get(LIST + "99/").status_code == 404


def test_保存_整えてページを作り直す(as_owner, site_repo):
    d = _read(site_repo)
    c = d["classroom"][1]
    c["reports"].insert(0, {"title": "新しい回", "date": "2030-01-05", "date_disp": "2030.1.5（土）開催",
                            "kind": "開催予定", "text": "こんにちは\n\n二段落目", "images": []})
    r = as_owner.post(LIST + "1/", _form(c, category="糠床作り教室", lead="ぬかどこ", action="save"), follow=True)
    page = r.content.decode()
    assert "保存しました。" in page and "ページを作り直しました（" in page and "件を掲載）。" in page
    cs = _read(site_repo)["classroom"]
    saved = cs[1]
    assert saved["category"] == "糠床作り教室" and saved["title"] == "糠床作り教室" and saved["lead"] == "ぬかどこ"
    # 回に番号が付く（既にあるものは変えない）。並べ替え用の日付はいちばん新しい回から
    ids = [x["id"] for x in saved["reports"]]
    assert ids[0] not in ("", None) and len(set(ids)) == len(ids)
    assert [x["id"] for x in c["reports"][1:]] == ids[1:]
    assert saved["date"] == "2030-01-05"
    assert saved["reports"][0]["text"] == "こんにちは\n\n二段落目"
    # ページができる（お教室と、その回）
    assert (site_repo / "classroom" / saved["id"] / "index.html").is_file()
    assert (site_repo / "classroom" / saved["id"] / ids[0] / "index.html").is_file()
    assert "糠床作り教室" in (site_repo / "classroom" / "index.html").read_text(encoding="utf-8")


def test_保存せずに足したり外したりできる(as_owner, site_repo):
    c = _read(site_repo)["classroom"][0]
    before = json.dumps(_read(site_repo), ensure_ascii=False)
    # 段落を足す → 入力欄が増えるが content.json は変わらない
    page = as_owner.post(LIST + "0/", _form(c, action="blockadd:list")).content.decode()
    assert 'name="blk_t_%d"' % len(c["blocks"]) in page and "1行が1項目になります。" in page
    assert "未保存の変更があります" in page
    assert json.dumps(_read(site_repo), ensure_ascii=False) == before
    # 開催予定を足すと、いちばん上に入る
    page = as_owner.post(LIST + "0/", _form(c, action="repadd:開催予定")).content.decode()
    assert "開催予定を作りました。日付とタイトルを入れてください。" in page
    assert 'name="rep_title_0" class="form-control" value=""' in page
    assert "%d回目" % (len(c["reports"]) + 1) in page
    page = as_owner.post(LIST + "0/", _form(c, action="repadd:過去の開催")).content.decode()
    assert "開催した回を作りました。日付とタイトルを入れてください。" in page
    # 新しい順に並べ替える（日付が無い回は下）
    c2 = dict(c, reports=[{"title": "古い", "date": "2020-01-01", "date_disp": "", "kind": "", "text": "", "images": []},
                          {"title": "日付なし", "date": "", "date_disp": "", "kind": "", "text": "", "images": []},
                          {"title": "新しい", "date": "2031-01-01", "date_disp": "", "kind": "", "text": "", "images": []}])
    page = as_owner.post(LIST + "0/", _form(c2, action="repsort")).content.decode()
    assert "新しい順に並べ替えました。「保存してページを作り直す」で確定します。" in page
    assert page.index('value="新しい"') < page.index('value="古い"') < page.index('value="日付なし"')
    # 回を消す・回を動かす
    page = as_owner.post(LIST + "0/", _form(c2, action="repdel:1")).content.decode()
    assert 'value="日付なし"' not in page and 'value="古い"' in page
    page = as_owner.post(LIST + "0/", _form(c2, action="repmove:0:down")).content.decode()
    assert page.index('value="日付なし"') < page.index('value="古い"')
    # 段落を消す
    page = as_owner.post(LIST + "0/", _form(c, action="blockdel:0")).content.decode()
    assert "まだ書かれていません。下のボタンで足してください。" in page
    # Enter で送られても（受け皿の action）何も壊れない
    page = as_owner.post(LIST + "0/", _form(c, action="keep")).content.decode()
    assert "お教室の編集：" in page
    assert json.dumps(_read(site_repo), ensure_ascii=False) == before


def test_写真_足す差し込む外す(as_owner, site_repo):
    c = _read(site_repo)["classroom"][0]
    P = _form(c, action="addimg")
    P["img_file"] = _png()
    page = as_owner.post(LIST + "0/", P).content.decode()
    assert "写真を追加しました。" in page
    saved = sorted(os.listdir(site_repo / "assets" / "img" / "classroom"))
    assert saved == ["classroom_photoone.webp"]        # 名前は英数字だけにして WebP に
    from PIL import Image

    im = Image.open(site_repo / "assets" / "img" / "classroom" / saved[0])
    assert im.width == 1200                           # 幅1200に縮める（旧アプリと同じ）
    assert 'name="images" value="classroom_photoone.webp"' in page and "本文に差し込む" in page
    assert "assets/img/classroom/classroom_photoone.webp" in page
    # 本文に差し込む → 写真の段落が増える
    c2 = dict(c, images=["classroom_photoone.webp"])
    page = as_owner.post(LIST + "0/", _form(c2, action="imgins:0")).content.decode()
    assert "概要の最後に差し込みました。" in page and "この位置から外す" in page
    assert 'name="blk_t_%d" value="img"' % len(c["blocks"]) in page
    # 外す → 写真も、差し込んだ段落も消える
    c3 = dict(c2, blocks=c["blocks"] + [{"t": "img", "v": "classroom_photoone.webp"}])
    page = as_owner.post(LIST + "0/", _form(c3, action="imgdel:0")).content.decode()
    assert "写真はまだありません。" in page and "この位置から外す" not in page
    # 開催した回の写真
    P = _form(c, action="repimg:0")
    P["rep_file_0"] = _png(300, 300)
    page = as_owner.post(LIST + "0/", P).content.decode()
    assert "写真を追加しました。" in page and 'name="rep_img_0" value="classroom_photoone_2.webp"' in page
    page = as_owner.post(LIST + "0/", _form(c, action="repimg:0")).content.decode()
    assert "失敗: ファイルが受け取れませんでした。" in page
    # 写真は手元のサイトから配られる
    r = as_owner.get("/manage/hp/site/assets/img/classroom/classroom_photoone.webp")
    assert r.status_code == 200
    # ここまで content.json には何も書いていない
    assert _read(site_repo)["classroom"][0]["images"] == c["images"]


def test_プレビュー_保存せずに見られる(as_owner, site_repo):
    c = _read(site_repo)["classroom"][0]
    page = as_owner.post(LIST + "0/", _form(c, lead="プレビューの説明", action="preview")).content.decode()
    assert "いま入力している内容で表示しています（保存はされていません）" in page
    assert "/manage/hp/preview/files/classroom/%s/" % c["id"] in page
    assert "スマホ" in page and "タブレット" in page and "入力を反映して更新" in page
    pv = site_repo / "admin" / "_preview" / "classroom" / c["id"] / "index.html"
    assert pv.is_file() and "プレビューの説明" in pv.read_text(encoding="utf-8")
    r = as_owner.get("/manage/hp/preview/files/classroom/%s/" % c["id"])
    assert r.status_code == 200
    assert _read(site_repo)["classroom"][0].get("lead", "") != "プレビューの説明"


def test_並べ替えと削除(as_owner, site_repo):
    before = [c["id"] for c in _read(site_repo)["classroom"]]
    r = as_owner.post(LIST, {"action": "move", "i": "0", "dir": "down", "next": LIST + "?cat="}, follow=True)
    assert "並べ替えました。" in r.content.decode()
    after = [c["id"] for c in _read(site_repo)["classroom"]]
    assert after[0] == before[1] and after[1] == before[0] and after[2:] == before[2:]
    # 端は動かない
    as_owner.post(LIST, {"action": "move", "i": "0", "dir": "up"})
    assert [c["id"] for c in _read(site_repo)["classroom"]] == after
    # 削除
    c = _read(site_repo)["classroom"][0]
    assert (site_repo / "classroom" / c["id"] / "index.html").is_file()
    r = as_owner.post(LIST + "0/", _form(c, action="del"), follow=True)
    assert "「%s」を削除しました。" % c["category"] in r.content.decode()
    assert [x["id"] for x in _read(site_repo)["classroom"]] == after[1:]
    assert not (site_repo / "classroom" / c["id"]).exists()


def test_サイトの表示が古いときに気づける(as_owner, site_repo):
    d = _read(site_repo)
    c = d["classroom"][0]
    # まず今の内容でページを作っておく
    as_owner.post(LIST + "0/", _form(c, action="save"))
    assert "サイトの表示が古くなっています" not in as_owner.get(LIST).content.decode()
    # 「開催予定」だった回の日付が過ぎた（content.json だけ変わり、HTMLは古いまま）
    d = _read(site_repo)
    d["classroom"][0]["reports"][0]["date"] = "2020-01-01"
    _write(site_repo, d)
    page = as_owner.get(LIST).content.decode()
    assert "サイトの表示が古くなっています" in page and c["category"] in page and "手元に反映" in page
    # 手元に反映 → 直る
    r = as_owner.post(LIST, {"action": "local"}, follow=True)
    assert "手元に反映" in r.content.decode()
    assert "サイトの表示が古くなっています" not in as_owner.get(LIST).content.decode()
