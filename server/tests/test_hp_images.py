"""公式サイト（apps/hp）: 写真。中身は mayumi-site/admin/app.py の sec_images と同じ。

写しの assets/img は空なので、試験の中で小さな画像を作って置く。
"""

import io
import json
import os

import pytest
from PIL import Image

pytestmark = pytest.mark.django_db

URL = "/manage/hp/images/"


def _画像(w=40, h=30, 種類="WEBP", 色=(200, 120, 80)):
    im = Image.new("RGB", (w, h), 色)
    b = io.BytesIO()
    im.save(b, 種類)
    return b.getvalue()


def _置く(site_repo, name, **kw):
    p = site_repo / "assets" / "img" / name
    p.write_bytes(_画像(**kw))
    return p


def _content(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get(URL)
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(URL).status_code in (302, 403)


def test_一覧_使われている写真と未使用の写真を見分ける(as_owner, site_repo):
    # head-outpatient.webp は母乳外来のページの見出し帯で使われている。nobody.webp はどこにも無い
    _置く(site_repo, "head-outpatient.webp")
    _置く(site_repo, "nobody.webp", w=800, h=600)
    # 入れ物のフォルダ（blog など）は一覧に入れない（消してはいけないため）
    (site_repo / "assets" / "img" / "blog").mkdir()
    (site_repo / "assets" / "img" / ".DS_Store").write_bytes(b"")
    page = as_owner.get(URL).content.decode()
    assert "🖼️ 写真の差し替え" in page and "🏞️ 見出し帯の写真" in page
    assert "head-outpatient.webp" in page and "nobody.webp" in page and "KB" in page
    assert page.count("未使用") == 1 + page.count("未使用の写真を消す")  # 札は nobody.webp だけ
    assert "どのページからも使われていない写真が 1 枚あります：nobody.webp" in page
    assert "未使用の写真を消す（1枚）" in page
    assert "使われていない写真 1 枚を削除します。よろしいですか？" in page
    assert 'value="blog"' not in page and ".DS_Store" not in page
    # 写真そのものは手元のサイトから配られる
    r = as_owner.get("/manage/hp/site/assets/img/nobody.webp")
    assert r.status_code == 200 and r["Content-Type"] == "image/webp"


def test_一覧_写真が無いとき(as_owner, site_repo):
    page = as_owner.get(URL).content.decode()
    assert "写真はまだありません" in page and "未使用の写真を消す" not in page


def test_差し替え_同じ名前で入れ替わり控えが残る(as_owner, site_repo):
    p = _置く(site_repo, "head-outpatient.webp", w=40, h=30)
    元 = p.read_bytes()
    新 = io.BytesIO(_画像(w=2000, h=1000, 種類="PNG", 色=(10, 20, 30)))
    新.name = "new.png"
    r = as_owner.post(URL, {"action": "replace", "target": "head-outpatient.webp", "file": 新}, follow=True)
    assert "写真を差し替えました: head-outpatient.webp" in r.content.decode()
    assert p.is_file() and p.read_bytes() != 元
    im = Image.open(p)
    assert im.format == "WEBP" and im.width == 1600  # 幅1600に縮めて WebP のまま
    assert (site_repo / "admin" / "backups" / "_images" / "head-outpatient.webp").read_bytes() == 元
    # 一覧に無い名前には入れない（新しい写真は「ページの編集」から）
    新.seek(0)
    r = as_owner.post(URL, {"action": "replace", "target": "../evil.webp", "file": 新}, follow=True)
    assert "差し替える写真が見つかりませんでした" in r.content.decode()
    assert not (site_repo / "assets" / "evil.webp").exists() and not (site_repo / "assets" / "img" / "evil.webp").exists()
    # 画像を選ばずに送る
    r = as_owner.post(URL, {"action": "replace", "target": "head-outpatient.webp"}, follow=True)
    assert "画像が選ばれていません" in r.content.decode()


def test_未使用の写真を消す_控えに移す(as_owner, site_repo):
    _置く(site_repo, "head-outpatient.webp")
    _置く(site_repo, "nobody.webp")
    _置く(site_repo, "nobody2.webp")
    r = as_owner.post(URL, {"action": "del_unused"}, follow=True)
    page = r.content.decode()
    assert "2枚を削除しました：nobody.webp、nobody2.webp" in page
    assert "（念のため backups/_images に移してあります）" in page
    img = site_repo / "assets" / "img"
    assert (img / "head-outpatient.webp").is_file()
    assert not (img / "nobody.webp").exists() and not (img / "nobody2.webp").exists()
    bak = site_repo / "admin" / "backups" / "_images"
    assert (bak / "nobody.webp").is_file() and (bak / "nobody2.webp").is_file()
    # もう一度押しても何も起きない
    r = as_owner.post(URL, {"action": "del_unused"}, follow=True)
    assert "未使用の写真はありませんでした。" in r.content.decode()


def test_見出し帯の写真_選んで下書き保存(as_owner, site_repo):
    _置く(site_repo, "head-outpatient.webp")
    _置く(site_repo, "nobody.webp")
    _置く(site_repo, "sns-line.png", 種類="PNG")
    page = as_owner.get(URL).content.decode()
    # いまの設定（写しの content.json）が出る
    assert 'name="ph_outpatient"' in page and 'value="head-outpatient.webp" selected' in page
    assert "（写真なし・色だけ）" in page and "膜の濃さ（0〜100）" in page
    assert 'name="ph_blog"' in page and "まゆみのつぶやき" in page
    # SNS のロゴは候補に出ない
    assert "sns-line.png" not in page.split("見出し帯の写真")[1].split("hp-imgs")[0]
    r = as_owner.post(URL, {
        "action": "save_heads", "ph_show": "0", "ph_veil": "150",
        "ph_outpatient": "nobody.webp", "ph_care02": "", "ph_blog": "存在しない.webp",
    }, follow=True)
    assert "下書きを保存しました（サイトにはまだ反映していません）。" in r.content.decode()
    c = _content(site_repo)["page_heads"]
    assert c["show"] is False and c["veil"] == 100
    assert c["images"]["outpatient"] == "nobody.webp" and c["images"]["care02"] == ""
    assert c["images"]["blog"] == ""  # 一覧に無い名前は入れない
    page = as_owner.get(URL).content.decode()
    assert 'value="0" selected' in page and 'value="100"' in page
    # 濃さに数字でないものを入れたら目安の 65 に戻る
    r = as_owner.post(URL, {"action": "save_heads", "ph_show": "1", "ph_veil": "abc"}, follow=True)
    c = _content(site_repo)["page_heads"]
    assert c["show"] is True and c["veil"] == 65
