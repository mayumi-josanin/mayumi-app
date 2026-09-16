"""公式サイト（apps/hp）: 選択肢の設定。中身は mayumi-site/admin/app.py の sec_options と同じ。"""

import json

import pytest

pytestmark = pytest.mark.django_db

URL = "/manage/hp/options/"


def _読む(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def _書く(site_repo, d):
    (site_repo / "admin" / "content.json").write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def _入力(site_repo):
    """画面に出ている入力を、いまの content.json から組み立てる（ブラウザが送る中身と同じ）。"""
    o = _読む(site_repo)["options"]
    post = {}
    for key in ("news_categories", "class_categories", "class_kinds"):
        for i, c in enumerate(o[key]):
            post["oc__%s__%d__name" % (key, i)] = c["name"]
            post["oc__%s__%d__color" % (key, i)] = c["color"]
    for i, w in enumerate(o["widths"]):
        post["oc__widths__%d__label" % i] = w["label"]
        post["oc__widths__%d__value" % i] = w["value"]
    return post


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get(URL)
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(URL).status_code in (302, 403)


def test_画面の見出しと選択肢と文言(as_owner, site_repo):
    page = as_owner.get(URL).content.decode()
    for 見出し in ("🎛️ 選択肢の設定", "📢 お知らせの種別", "🎒 お教室の正式な名称",
                "🗂 お教室の分類（定期開催・臨時開催）", "📐 「ページの編集」の幅", "ℹ️ ここで変えられないもの"):
        assert 見出し in page, 見出し
    # いまの選択肢（content.json の写し）が並ぶ
    o = _読む(site_repo)["options"]
    for c in o["news_categories"] + o["class_categories"] + o["class_kinds"]:
        assert 'value="%s"' % c["name"] in page and 'value="%s"' % c["color"] in page
    for w in o["widths"]:
        assert 'value="%s"' % w["label"] in page
    # ボタンとヒントの文言は旧アプリと同じ
    for 文 in ("＋ 種別を追加", "＋ お教室の名称を追加", "＋ 分類を追加", "＋ 幅を追加", "削除", "↑", "↓",
              "名前を変えると、すでにその種別を付けたお知らせは「なし」扱いになります",
              "画面に出す名前", "40% など（空＝そのまま）", "仕組みと結びついているため設定では変えられません",
              "を選択肢から消します。よろしいですか？"):
        assert 文 in page, 文
    # 左メニューにも出る（診療カレンダーの次・プレビューの前）
    assert 'class="active">選択肢</a>' in page  # 上部タブ（左メニューは「ページを整える」にまとめた）


def test_保存_名前と色と幅が_content_json_に入る(as_owner, site_repo):
    post = _入力(site_repo)
    post["oc__news_categories__0__name"] = "イベントのお知らせ"
    post["oc__news_categories__0__color"] = "#123456"
    post["oc__class_kinds__1__name"] = "特別開催"
    post["oc__widths__1__label"] = "小さめ 30%"
    post["oc__widths__1__value"] = "30%"
    post["action"] = "save"
    r = as_owner.post(URL, post, follow=True)
    assert "下書きを保存しました（サイトにはまだ反映していません）。" in r.content.decode()
    o = _読む(site_repo)["options"]
    assert o["news_categories"][0] == {"name": "イベントのお知らせ", "color": "#123456"}
    assert o["class_kinds"][1]["name"] == "特別開催"
    assert o["widths"][1] == {"label": "小さめ 30%", "value": "30%"}
    # ほかは変わらない
    assert o["class_categories"][0]["name"] == "離乳食教室"


def test_追加_4種類とも初期値で増える(as_owner, site_repo):
    前 = _読む(site_repo)["options"]
    期待 = {
        "news_categories": {"name": "新しい種別", "color": "#8a9e7e"},
        "class_categories": {"name": "新しいお教室", "color": "#8a9e7e"},
        "class_kinds": {"name": "新しい分類", "color": "#8a9e7e"},
        "widths": {"label": "新しい幅", "value": "50%"},
    }
    for key, 中身 in 期待.items():
        post = _入力(site_repo)
        post["action"] = "add|" + key
        as_owner.post(URL, post)
        o = _読む(site_repo)["options"]
        assert len(o[key]) == len(前[key]) + 1 and o[key][-1] == 中身, key
    page = as_owner.get(URL).content.decode()
    assert 'value="新しい種別"' in page and 'value="新しい幅"' in page


def test_削除_その番号だけ消える(as_owner, site_repo):
    前 = _読む(site_repo)["options"]["class_categories"]
    post = _入力(site_repo)
    post["action"] = "del|class_categories|1"  # 糠床作り教室
    as_owner.post(URL, post)
    後 = _読む(site_repo)["options"]["class_categories"]
    assert [c["name"] for c in 後] == [c["name"] for c in 前 if c["name"] != "糠床作り教室"]
    # 使われている選択肢でも止めない（旧アプリと同じ。確認は画面の confirm だけで、
    # 消したあと付いていたお教室は「なし」扱いになる）


def test_並べ替え_上へ下へ_端では動かない(as_owner, site_repo):
    前 = [c["name"] for c in _読む(site_repo)["options"]["news_categories"]]
    post = _入力(site_repo); post["action"] = "move|news_categories|1|-1"
    as_owner.post(URL, post)
    assert [c["name"] for c in _読む(site_repo)["options"]["news_categories"]] == [前[1], 前[0]] + 前[2:]
    post = _入力(site_repo); post["action"] = "move|news_categories|0|1"
    as_owner.post(URL, post)
    assert [c["name"] for c in _読む(site_repo)["options"]["news_categories"]] == 前
    # 先頭を上へ・末尾を下へ は何も起きない
    post = _入力(site_repo); post["action"] = "move|news_categories|0|-1"
    as_owner.post(URL, post)
    post = _入力(site_repo); post["action"] = "move|widths|%d|1" % (len(_読む(site_repo)["options"]["widths"]) - 1)
    as_owner.post(URL, post)
    assert [c["name"] for c in _読む(site_repo)["options"]["news_categories"]] == 前


def test_書きかけの名前は_追加や並べ替えでも消えない(as_owner, site_repo):
    """旧アプリは操作の前に画面の入力を集めていた（collect）。同じにする。"""
    post = _入力(site_repo)
    post["oc__news_categories__2__name"] = "商品のお知らせ"
    post["action"] = "add|news_categories"
    as_owner.post(URL, post)
    o = _読む(site_repo)["options"]["news_categories"]
    assert o[2]["name"] == "商品のお知らせ" and o[-1]["name"] == "新しい種別"


def test_選択肢が入っていない古いデータでも既定値が出て直せる(as_owner, site_repo):
    d = _読む(site_repo)
    d.pop("options", None)
    _書く(site_repo, d)
    page = as_owner.get(URL).content.decode()
    assert 'value="イベント情報"' in page and 'value="定期開催"' in page and 'value="いっぱい"' in page
    post = {"oc__news_categories__0__name": "イベント案内", "action": "save"}
    as_owner.post(URL, post)
    o = _読む(site_repo)["options"]
    assert o["news_categories"][0]["name"] == "イベント案内" and len(o["widths"]) == 5


def test_おかしな操作は無視して落ちない(as_owner, site_repo):
    前 = _読む(site_repo)["options"]
    for action in ("del|news_categories|99", "move|widths|x|1", "add|nothing", "del|", "なにか"):
        post = _入力(site_repo); post["action"] = action
        assert as_owner.post(URL, post).status_code == 302, action
    assert _読む(site_repo)["options"] == 前
