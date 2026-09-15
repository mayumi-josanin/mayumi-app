"""公式サイト（apps/hp）: 📅 診療カレンダー。中身は mayumi-site/admin/app.py の sec_closed と同じ。"""

import io
import re
import json
from datetime import date

import pytest

from apps.content.models import CalendarEvent

pytestmark = pytest.mark.django_db

URL = "/manage/hp/closed/"


def _content(site_repo):
    return json.loads((site_repo / "admin" / "content.json").read_text(encoding="utf-8"))


def _write(site_repo, d):
    (site_repo / "admin" / "content.json").write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (1400, 900), (200, 120, 80)).save(buf, "PNG")
    buf.seek(0)
    buf.name = "omatsuri.png"
    return buf


def _event(row, on, title, category="", color="", **extra):
    kw = dict(sheet_row=row, event_on=date.fromisoformat(on), title=title, category=category,
              color=color, published=True, notice_listed=True)
    kw.update(extra)
    return CalendarEvent.objects.create(**kw)


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    r = as_owner.get(URL)
    assert r.status_code == 200 and "まだ設定されていません" in r.content.decode()


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get(URL).status_code in (302, 403)


def test_画面_件数と一覧と絞り込み(as_owner, site_repo):
    page = as_owner.get(URL).content.decode()
    # 本物の写しには 休診(sc01)・往診(sc03)・訪問産後ケア(sc04)・空の予定(sc02) がある
    assert "📅 診療カレンダー" in page and "＋ 新しい予定" in page
    assert "📥 アプリのカレンダーから取り込む" in page and "サイトのカレンダーを開く" in page
    assert "休診 69件" in page and "往診 18件" in page and "訪問産後ケア 5件" in page
    assert "全4件" in page and "（名前なし）" in page and "日付なし" in page
    assert "2026-03-26、2026-04-01、2026-04-05 ほか66日" in page and "69日" in page
    assert "毎週の定休日（水・日・祝）は入れなくて大丈夫です" in page
    assert "自動で入る予定" in page and "お知らせは、以前は自動で入っていましたが" in page
    # 種類でしぼる
    page = as_owner.get(URL + "?kind=visit").content.decode()
    assert "1件 / 全4件" in page and "ほか15日" in page
    page = as_owner.get(URL + "?kind=event").content.decode()
    assert "その種類の予定はありません。" in page
    # 予定が1つも無いとき
    d = _content(site_repo)
    d["schedule"] = []
    _write(site_repo, d)
    page = as_owner.get(URL).content.decode()
    assert "まだありません。「＋ 新しい予定」から作れます。" in page and "休診 0件" in page


def test_カレンダーで見る_月と日(as_owner, site_repo):
    page = as_owner.get(URL + "?ym=2026-09&day=2026-09-08").content.decode()
    assert "2026年 9月" in page and "前の月" in page and "今月" in page and "次の月" in page
    assert "9月8日（火）" in page and 'class="calv-auto"' not in page
    # 休診(sc01) は一覧の2番目なので編集の行き先は edit=1
    assert "&amp;edit=1#scEdit" in page or "&edit=1#scEdit" in page
    page = as_owner.get(URL + "?ym=2026-09&day=2026-09-07").content.decode()
    assert "9月7日（月）" in page and "この日の予定はありません。" in page
    # 変な月は今月に戻す
    assert as_owner.get(URL + "?ym=2026-13").status_code == 200


def test_新しい予定_保存_削除(as_owner, site_repo):
    first = as_owner.get(URL).content.decode()
    r = as_owner.post(URL, {"action": "new"}, follow=True)
    page = r.content.decode()
    assert r.redirect_chain[-1][0].endswith("?edit=0&new=1")
    assert "新しい予定" in page and "日付を選んで「追加」、名前を入れて保存してください。" in page
    assert "日付を選ぶ" in page and "選択中の日付" in page and "保存してカレンダーに反映" in page
    d = _content(site_repo)
    assert d["schedule"][0]["id"] == "sc05" and d["schedule"][0]["kind"] == "closed"

    # 日付が無ければ保存しない（入力は画面に残る）
    r = as_owner.post(URL, {"action": "save", "index": "0", "kind": "event", "title": "夏祭り",
                            "show": "1", "note": "午後のみ", "text": ""})
    page = r.content.decode()
    assert "日付が1つも入っていません。" in page and 'value="夏祭り"' in page
    assert _content(site_repo)["schedule"][0]["title"] == ""

    # 日付・名前・本文・写真を入れて保存 → カレンダーと予定のページを作り直す
    r = as_owner.post(URL, {"action": "save", "index": "0", "kind": "event", "title": "夏祭り",
                            "show": "1", "note": "午後のみ", "text": "みんなで踊ります。\n\n**雨天中止**",
                            "dates": ["2026-08-20", "2026-08-19", "2026-08-20"], "photo": _png()}, follow=True)
    page = r.content.decode()
    assert "保存しました。" in page and "詳しいページ 1件" in page and "写真を追加しました（calendar_omatsuri.webp）" in page
    assert "予定の編集：夏祭り" in page
    c = _content(site_repo)["schedule"][0]
    assert c["dates"] == ["2026-08-19", "2026-08-20"] and c["note"] == "午後のみ" and c["images"] == ["calendar_omatsuri.webp"]
    assert (site_repo / "assets" / "img" / "calendar" / "calendar_omatsuri.webp").is_file()
    assert (site_repo / "calendar" / "sc05" / "index.html").is_file()
    assert "夏祭り" in (site_repo / "reception" / "index.html").read_text(encoding="utf-8")
    # お教室の開催日も「イベント」に数えるので、2日ぶん増えたことで見る
    before = int(re.search(r"イベント (\d+)件", first).group(1))
    assert "イベント %d件" % (before + 2) in page and "くわしい内容あり" in page
    # 写真が画面に出て「外す」で外せる（hidden の images を送らない）
    assert "assets/img/calendar/calendar_omatsuri.webp" in page
    r = as_owner.post(URL, {"action": "save", "index": "0", "kind": "event", "title": "夏祭り", "show": "0",
                            "note": "", "text": "", "dates": ["2026-08-19"]}, follow=True)
    c = _content(site_repo)["schedule"][0]
    assert c["images"] == [] and c["show"] is False
    assert not (site_repo / "calendar" / "sc05").exists()      # 中身が無くなったのでページは片づく
    assert "隠す" in r.content.decode()

    # 削除
    r = as_owner.post(URL, {"action": "del", "index": "0"}, follow=True)
    assert "「夏祭り」を削除しました。" in r.content.decode()
    assert len(_content(site_repo)["schedule"]) == 4
    r = as_owner.post(URL, {"action": "del", "index": "99"}, follow=True)
    assert "その位置の予定がありません。" in r.content.decode()


def test_アプリのカレンダーから取り込む(as_owner, site_repo):
    _event(2, "2026-10-01", "休診日", category="休診")
    _event(3, "2026-10-02", "往診", color="#ffc7fa")
    _event(4, "2026-10-03", "訪問型産後ケア", color="#e57373")   # 色は休診だが題名を先に見る
    _event(5, "2026-10-04", "ベビーマッサージ教室", category="イベント")  # イベントは持ってこない
    _event(6, "2026-10-05", "休診", category="休診", published=False)   # 非公開は出さない
    _event(7, "2026-10-06", "休診", category="休診", deleted=True)      # 消したものは出さない
    _event(8, "2026-09-08", "休診", category="休診")                    # もう入っている
    _event(9, "2026-10-01", "休診", category="休診")                    # 同じ種類・同じ日は1つに
    page = as_owner.get(URL + "?import=1").content.decode()
    assert "お客様アプリのカレンダーにある<b>休診・往診・訪問産後ケア</b>を持ってきます。" in page
    assert "すべて選ぶ" in page and "今日以降のものだけ" in page and "まだ入れていないものだけ" in page
    assert 'value="closed|2026-10-01"' in page and 'value="visit|2026-10-02"' in page and 'value="postpartum|2026-10-03"' in page
    assert "2026-10-04" not in page and "2026-10-05" not in page and "2026-10-06" not in page
    assert page.count('value="closed|2026-10-01"') == 1
    assert 'value="closed|2026-09-08" disabled' in page and "入っています" in page
    assert "訪問型産後ケア" in page          # 題名が種類の名前と違うときは添える

    r = as_owner.post(URL, {"action": "import"}, follow=True)
    assert "取り込む日が選ばれていません。" in r.content.decode()

    r = as_owner.post(URL, {"action": "import", "keys": ["closed|2026-10-01", "visit|2026-10-02", "postpartum|2026-10-03", "closed|2026-09-08"]}, follow=True)
    page = r.content.decode()
    assert "3日を取り込みました。" in page
    assert "休診 に 1日を足しました（ぜんぶで 70日）。" in page
    assert "往診 に 1日を足しました（ぜんぶで 19日）。" in page
    assert "訪問産後ケア に 1日を足しました（ぜんぶで 6日）。" in page
    d = _content(site_repo)
    by = {c["id"]: c for c in d["schedule"]}
    assert "2026-10-01" in by["sc01"]["dates"] and "2026-10-02" in by["sc03"]["dates"] and "2026-10-03" in by["sc04"]["dates"]
    # 同じ種類の予定が無ければ作る
    d["schedule"] = [c for c in d["schedule"] if c["kind"] != "visit"]
    _write(site_repo, d)
    _event(10, "2026-10-09", "往診", category="往診")
    r = as_owner.post(URL, {"action": "import", "keys": ["visit|2026-10-09", "visit|2026-10-02"]}, follow=True)
    page = r.content.decode()
    assert "「往診」の予定を作りました。" in page and "往診 に 2日を足しました（ぜんぶで 2日）。" in page
    assert [c for c in _content(site_repo)["schedule"] if c["kind"] == "visit"][0]["dates"] == ["2026-10-02", "2026-10-09"]
    # もう全部入っていれば
    page = as_owner.get(URL + "?import=1").content.decode()
    assert "アプリのカレンダーの休診・往診・訪問産後ケアは、すべて入っています。" in page
    r = as_owner.post(URL, {"action": "import", "keys": ["closed|2026-10-01"]}, follow=True)
    assert "取り込むものがありませんでした。" in r.content.decode()
