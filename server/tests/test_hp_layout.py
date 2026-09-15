"""公式サイト（apps/hp）: 🧩 ページの編集。中身は mayumi-site/admin/app.py の sec_layout / api/layout-* と同じ。

小さなページ（access・staff・outpatient）で、かたまりの取り出し・文章の変更・写真の追加・
見え方（プレビュー）・保存を固定する。本物の mayumi-site には何も書かない（site_repo は写し）。
"""

import io
import json
import os

import pytest

pytestmark = pytest.mark.django_db

URL = "/manage/hp/layout/"


def _data(client, page):
    r = client.get(URL + "data/", {"page": page})
    assert r.status_code == 200
    j = r.json()
    assert j["ok"], j
    return j["data"]


def _post(client, what, body):
    r = client.post(URL + "api/%s/" % what, data=json.dumps(body), content_type="application/json")
    assert r.status_code == 200, (what, r.status_code)
    return r.json()


def _png():
    from PIL import Image

    im = Image.new("RGB", (40, 30), (200, 120, 80))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get(URL)
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()
    assert as_owner.get(URL + "data/", {"page": "access"}).json()["ok"] is False
    assert _post(as_owner, "runs", {"blocks": []})["ok"] is False


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(URL).status_code in (302, 403)
    assert client.get(URL + "data/", {"page": "access"}).status_code in (302, 403)


def test_画面_ページを選ぶ欄と旧アプリの文言(as_owner, site_repo):
    page = as_owner.get(URL).content.decode()
    for s in ("ページの編集", "母乳外来", "産後ケア", "アクセス", "読み直す", "やり直す", "この配置で確定",
              "つまんで上下にドラッグ", "見え方", "端末の画面サイズ", "ここに追加する", "この写真を追加",
              "文章として追加", "見出し（大）として追加"):
        assert s in page, s
    # 選んだページが選ばれた状態で開く
    page = as_owner.get(URL, {"page": "staff"}).content.decode()
    assert 'value="staff" selected' in page
    # 無いページはいちばん上のページになる
    page = as_owner.get(URL, {"page": "nope"}).content.decode()
    assert 'value="outpatient" selected' in page


def test_かたまりの取り出し(as_owner, site_repo):
    d = _data(as_owner, "access")
    assert d["blocks"] and d["runs"] and d["widths"] and d["aligns"]
    b = d["blocks"][0]
    for k in ("label", "summary", "opts", "sizable", "editable", "text", "parts", "lists", "price_key"):
        assert k in b, k
    assert any(b["editable"] == "text" for b in d["blocks"])
    assert [w[0] for w in d["widths"]][:2] == ["そのまま", "小 40%"]
    assert d["aligns"][0] == ["中央", "center"]
    assert "bonyu-massage" in d["prices"] and d["price_labels"]["itothermy"] == "イトオテルミー温熱療法"
    # 無いページ
    r = as_owner.get(URL + "data/", {"page": "index"}).json()
    assert r["ok"] is False and "そのページはありません" in r["log"]


def test_文章を直す_まるごと(as_owner, site_repo):
    d = _data(as_owner, "access")
    i, b = next((i, b) for i, b in enumerate(d["blocks"]) if b["editable"] == "text" and b["tag"] == "p")
    r = _post(as_owner, "text", {"block": b, "text": "試験で直した文章です。\n**太字**もあります。"})
    assert r["ok"], r
    nb = r["block"]
    assert "試験で直した文章です。" in nb["html"]
    assert "<b>太字</b>" in nb["html"] or "<strong>太字</strong>" in nb["html"]
    assert nb["text"].startswith("試験で直した文章です。") and nb["label"] == "文章"
    assert nb["summary"].startswith("試験で直した文章です。")


def test_本文の枠_ひと続きの文章を1つの枠で直す(as_owner, site_repo):
    d = _data(as_owner, "access")
    run = d["runs"][1]
    assert run["z"] - run["a"] >= 2 and "\n" in run["text"]
    text = run["text"] + "\n## 試験で足した中見出し\n足した本文です。\n- 足した項目1\n- 足した項目2"
    r = _post(as_owner, "run", {"a": run["a"], "z": run["z"], "blocks": d["blocks"], "text": text})
    assert r["ok"], r
    assert len(r["blocks"]) == len(d["blocks"]) + 3  # h3 + p + ul
    tags = [b["tag"] for b in r["blocks"][run["a"]:run["z"] + 3]]
    assert tags[-3:] == ["h3", "p", "ul"]
    assert "足した項目2" in r["blocks"][run["z"] + 2]["html"]
    # 枠の範囲も計算し直されている
    assert any(x["a"] == run["a"] and x["z"] == run["z"] + 3 for x in r["runs"])
    # 前の部分は1文字も変わらない
    assert [b["html"] for b in r["blocks"][:run["a"]]] == [b["html"] for b in d["blocks"][:run["a"]]]


def test_並べ替え_枠の範囲を計算し直す(as_owner, site_repo):
    d = _data(as_owner, "access")
    blocks = d["blocks"]
    moved = blocks[1:] + blocks[:1]   # 先頭のかたまり（写真）を末尾へ
    r = _post(as_owner, "runs", {"blocks": moved})
    assert r["ok"] and r["runs"] and r["runs"][0]["a"] == 0
    assert r["runs"] != d["runs"]


def test_幅と寄せ(as_owner, site_repo):
    d = _data(as_owner, "care02")
    i, b = next((i, b) for i, b in enumerate(d["blocks"]) if b["sizable"] and b["tag"] == "img")
    r = _post(as_owner, "opt", {"page": "care02", "block": b, "key": "width", "value": "60%"})
    assert r["ok"] and r["block"]["opts"]["width"] == "60%" and "60%" in r["block"]["html"]
    r = _post(as_owner, "opt", {"page": "care02", "block": r["block"], "key": "align", "value": "left"})
    assert r["ok"] and r["block"]["opts"]["align"] == "left" and r["block"]["opts"]["width"] == "60%"


def test_追加_文章と見出し(as_owner, site_repo):
    r = _post(as_owner, "new", {"kind": "p", "value": "新しい文章"})
    assert r["ok"], r
    assert r["block"]["tag"] == "p" and "新しい文章" in r["block"]["html"]
    assert r["block"]["editable"] == "text" and r["block"]["label"] == "文章"
    r = _post(as_owner, "new", {"kind": "h2", "value": ""})
    assert r["ok"] and r["block"]["tag"] == "h2" and "見出し" in r["block"]["html"]
    r = _post(as_owner, "new", {"kind": "div", "value": ""})
    assert r["ok"] is False and "不明な種類" in r["log"]


def test_かたまりの中の文章だけを直す(as_owner, site_repo):
    d = _data(as_owner, "staff")
    b = next(b for b in d["blocks"] if b["parts"] and b["lists"])
    # 写真の説明（空）ではなく、文字のある欄（役職）を直す
    p0 = next(p for p in b["parts"] if p["text"].strip())
    r = _post(as_owner, "parts", {"block": b, "values": {str(p0["i"]): "試験で直した役職"}, "lists": {}})
    assert r["ok"], r
    assert "試験で直した役職" in r["block"]["html"] and p0["text"] not in r["block"]["html"]
    # 直したところ以外の中身は同じ
    other = [p["text"] for p in b["parts"] if p["i"] != p0["i"]]
    assert [p["text"] for p in r["block"]["parts"] if p["i"] != p0["i"]] == other
    # 箇条書きは行の数だけ項目になる
    L = b["lists"][0]
    r = _post(as_owner, "parts", {"block": b, "values": {}, "lists": {str(L["i"]): "項目い\n項目ろ\n項目は"}})
    assert r["ok"], r
    assert r["block"]["lists"][0]["n"] == 3 and "項目ろ" in r["block"]["html"]


def test_見え方_プレビューを作って管理画面の中で開く(as_owner, site_repo):
    d = _data(as_owner, "access")
    i, b = next((i, b) for i, b in enumerate(d["blocks"]) if b["editable"] == "text")
    d["blocks"][i] = _post(as_owner, "text", {"block": b, "text": "見え方だけの文章"})["block"]
    r = _post(as_owner, "preview", {"page": "access", "data": d})
    assert r["ok"], r
    pv = site_repo / "admin" / "_preview" / "access" / "index.html"
    assert pv.is_file()
    html = pv.read_text(encoding="utf-8")
    assert "見え方だけの文章" in html and 'data-ed="' in html
    # サイトのファイルはまだ変わらない
    assert "見え方だけの文章" not in (site_repo / "access" / "index.html").read_text(encoding="utf-8")
    # 管理画面の中で見る。文字を押して直す仕掛けが差し込まれ、リンクはプレビューの住所になる
    r = as_owner.get(URL + "pv/access/")
    assert r.status_code == 200
    out = r.content.decode()
    assert "ed-preview" in out and '<base href="/manage/hp/preview/files/access/">' in out
    assert 'href="/manage/hp/preview/files/' in out and 'href="/care02/"' not in out
    assert "見え方だけの文章" in out
    # 差し込んだ仕掛けはファイルには残らない
    assert "ed-preview" not in html
    assert as_owner.get(URL + "pv/index/").status_code == 404
    r = _post(as_owner, "preview", {"page": "index", "data": d})
    assert r["ok"] is False


def test_この配置で確定_ファイルに書き込み控えを取る(as_owner, site_repo):
    d = _data(as_owner, "access")
    i, b = next((i, b) for i, b in enumerate(d["blocks"]) if b["editable"] == "text" and b["tag"] == "p")
    d["blocks"][i] = _post(as_owner, "text", {"block": b, "text": "確定した文章です。"})["block"]
    # 末尾に文章を1つ足す
    d["blocks"].append(_post(as_owner, "new", {"kind": "p", "value": "足した文章です。"})["block"])
    r = _post(as_owner, "save", {"page": "access", "data": d})
    assert r["ok"] and r["log"][0].startswith("バックアップ: ") and r["log"][1] == "配置を書き込みました。"
    page = (site_repo / "access" / "index.html").read_text(encoding="utf-8")
    assert "確定した文章です。" in page and "足した文章です。" in page and b["html"] not in page
    assert os.listdir(site_repo / "admin" / "backups")
    # 読み直すと直した内容で取り出せる
    d2 = _data(as_owner, "access")
    assert d2["blocks"][i]["text"] == "確定した文章です。" and d2["blocks"][-1]["text"] == "足した文章です。"
    assert len(d2["blocks"]) == len(d["blocks"])
    # 同じ内容でもう一度確定しても、変更なし
    r = _post(as_owner, "save", {"page": "access", "data": d2})
    assert r["log"][1] == "変更はありませんでした。"
    r = _post(as_owner, "save", {"page": "index", "data": d2})
    assert r["ok"] is False


def test_写真を足す(as_owner, site_repo):
    r = as_owner.post(URL + "new-image/", {"file": io.BytesIO(_png()), "alt": "試験の写真"})
    assert r.status_code == 200
    j = r.json()
    assert j["ok"], j
    assert j["filename"].endswith(".webp") and (site_repo / "assets" / "img" / j["filename"]).is_file()
    b = j["block"]
    assert b["tag"] == "img" and b["sizable"] and b["label"] == "写真" and 'alt="試験の写真"' in b["html"]
    assert b["opts"]["width"] == "" and b["opts"]["align"] == "center"
    r = as_owner.post(URL + "new-image/", {"alt": "x"})
    assert r.json()["ok"] is False and "ファイルが受け取れませんでした" in r.json()["log"]


def test_料金表を直す(as_owner, site_repo):
    d = _data(as_owner, "outpatient")
    i, b = next((i, b) for i, b in enumerate(d["blocks"])
                if b["price_key"] and not d["prices"][b["price_key"]].get("merge"))
    key = b["price_key"]
    p = d["prices"][key]
    rows = [list(r) for r in p["rows"]] + [["試験の項目", "9,999円"]]
    r = _post(as_owner, "price-save", {"key": key, "block": b, "caption": "試験の見出し", "note": "※試験", "rows": rows})
    assert r["ok"], r
    assert "試験の見出し" in r["block"]["html"] and "試験の項目" in r["block"]["html"]
    saved = json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))["prices"][key]
    assert saved["caption"] == "試験の見出し" and saved["note"] == "※試験" and saved["rows"][-1] == ["試験の項目", "9,999円"]
    assert r["prices"][key]["caption"] == "試験の見出し"
    # 👁 プレビューは content.json を変えない
    r = _post(as_owner, "price-preview", {"key": key, "block": b, "caption": "見るだけ", "note": "", "rows": rows})
    assert r["ok"] and "見るだけ" in r["block"]["html"]
    assert json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))["prices"][key]["caption"] == "試験の見出し"
    # 見え方の画面のマス目から直す
    r = _post(as_owner, "price-cell", {"key": key, "block": b, "r": 0, "c": 1, "text": "1円"})
    assert r["ok"], r
    assert json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))["prices"][key]["rows"][0][1] == "1円"
    r = _post(as_owner, "price-cell", {"key": key, "block": b, "r": 99, "c": 1, "text": "1円"})
    assert r["ok"] is False
    r = _post(as_owner, "price-save", {"key": "nope", "block": b, "caption": "", "note": "", "rows": []})
    assert r["ok"] is False and "その料金表はありません" in r["log"]
