"""NEWS配信・管理。**中身は旧管理アプリ（NEWS配信・管理）と同じ**
（入力項目・分類とカテゴリの絞り込み・新しい順・公開状況の select・一括削除・通知の文言）。
読み書きは gasapi/admin_news.py（GAS の転送先と同じ）。"""

import io

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile
from django.utils import timezone

from apps.content.models import Category, News
from apps.gasapi.admin_news import 一覧

pytestmark = pytest.mark.django_db


@pytest.fixture
def categories(db):
    Category.objects.create(sheet_row=2, name="イベント情報", kind="お知らせ")
    Category.objects.create(sheet_row=3, name="まゆみのつぶやき", kind="ブログ")
    Category.objects.create(sheet_row=4, name="よもぎ茶", kind="商品")  # NEWS の候補には出ない


@pytest.fixture
def news(categories):
    return News.objects.create(
        sheet_row=10, posted_on=timezone.localdate(), title="米味噌作り", category="イベント情報",
        icon="📢", body="本文", published=True, notice_listed=True,
        image_url='["https://drive.google.com/thumbnail?id=abc"]',
    )


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (2400, 1200), (200, 120, 80)).save(buf, format="PNG")
    return SimpleUploadedFile("photo.png", buf.getvalue(), content_type="image/png")


def _form(**extra):
    data = {
        "date": "2026-09-14T10:30", "title": "マッサージ教室 おさらい会", "category": "イベント情報",
        "icon": "📢", "body": "9月14日 10:00〜11:30", "publishAt": "", "linkUrl": "",
        "linkButtonText": "", "notice_refresh": "on", "status": "公開",
    }
    data.update(extra)
    return data


# ========== 一覧 ==========


def test_一覧の見出しと列と空の文言(as_owner, categories):
    page = as_owner.get("/manage/news/").content.decode()
    for s in ["NEWS配信・管理", "下書き保存一覧", "投稿済み記事一覧", "🗑 一括削除", "🔄 更新",
              "全ての分類", "全てのカテゴリ", "下書き保存はありません", "投稿済み記事はありません"]:
        assert s in page, s
    for th in ["<th>アイコン</th>", "<th>日付</th>", "<th>カテゴリ</th>", "<th>分類</th>", "<th>タイトル</th>", "<th>公開状況</th>", "<th>操作</th>"]:
        assert th in page, th
    # カテゴリの絞り込みには NEWS 用（お知らせ／ブログ）だけ
    assert 'value="イベント情報"' in page and 'value="まゆみのつぶやき"' in page and 'value="よもぎ茶"' not in page


def test_一覧に投稿済みと下書きが分かれて出る(as_owner, news):
    News.objects.create(sheet_row=11, posted_on=timezone.localdate(), title="下書きの記事",
                        category="まゆみのつぶやき", published=False)
    News.objects.create(sheet_row=12, posted_on=timezone.localdate(), title="消した記事",
                        category="イベント情報", published=True, deleted=True)
    page = as_owner.get("/manage/news/").content.decode()
    assert "米味噌作り" in page and "下書きの記事" in page and "消した記事" not in page
    # 分類はカテゴリ側が持つ（お知らせ＝赤、ブログ＝緑）。行内の公開状況は select
    assert "badge-red" in page and "badge-green" in page
    assert page.count('name="status"') == 2 and 'value="非公開" selected' in page
    assert page.index("下書き保存一覧") < page.index("投稿済み記事一覧")


def test_新しい順に並ぶ(as_owner, categories):
    import datetime

    News.objects.create(sheet_row=5, posted_on=datetime.date(2026, 9, 1), title="古い記事", category="イベント情報", published=True)
    News.objects.create(sheet_row=6, posted_on=datetime.date(2026, 9, 10), title="同じ日の先", category="イベント情報", published=True)
    News.objects.create(sheet_row=7, posted_on=datetime.date(2026, 9, 10), title="同じ日の後", category="イベント情報", published=True)
    page = as_owner.get("/manage/news/").content.decode()
    assert page.index("同じ日の後") < page.index("同じ日の先") < page.index("古い記事")


def test_分類とカテゴリで絞れる(as_owner, news):
    News.objects.create(sheet_row=11, posted_on=timezone.localdate(), title="つぶやき",
                        category="まゆみのつぶやき", published=True)
    News.objects.create(sheet_row=12, posted_on=timezone.localdate(), title="つぶやきの下書き",
                        category="まゆみのつぶやき", published=False)
    page = as_owner.get("/manage/news/?kind=ブログ").content.decode()
    assert "つぶやき" in page and "米味噌作り" not in page
    page = as_owner.get("/manage/news/?category=イベント情報").content.decode()
    # 絞り込みは下書きにも効く
    assert "米味噌作り" in page and "つぶやきの下書き" not in page and "下書き保存はありません" in page


# ========== 投稿 ==========


def test_投稿画面の項目と文言(as_owner, categories):
    page = as_owner.get("/manage/news/new/").content.decode()
    for s in ["日時", "タイトル", "例：夏の産後ヨガ体験会", "★画像をアップロード（スマホのライブラリから選択可能）",
              "※画像全体を保ったまま、アップロード用に最適化して取り込みます", "カテゴリ", "アイコン", "本文",
              "選択した文字に太字・下線を設定できます", "お知らせの内容を入力...", "公開開始日時（空欄ならすぐ公開）",
              "リンクURL（任意：詳しく見るボタンの遷移先）", "※入力すると、アプリ側でボタンが表示されます。",
              "ボタンのテキスト（任意）", "詳しく見る （未入力ならこれになります）", "※「予約はこちら」「インスタを見る」など自由に入力できます。",
              "この保存をお知らせ一覧に反映する", "初投稿はオンのまま、編集時は必要なときだけチェックしてください。",
              "新規投稿は既定でオンです。", "この保存でPush通知を送信する", "送信したいときだけ、右側のチェックを入れて保存してください。",
              "未チェックでは通知されません。", "👁 プレビュー", "🌟 公開して投稿する", "📁 下書き保存",
              '<optgroup label="お知らせ分類">', '<optgroup label="ブログ分類">']:
        assert s in page, s
    assert 'name="notice_refresh" aria-label="NEWS保存時にお知らせ一覧へ反映" checked' in page
    assert 'value="📢"' in page  # アイコンの既定
    assert 'value="よもぎ茶"' not in page  # 商品のカテゴリは出ない


def test_カテゴリ管理にNEWSの分類が無ければ既定の候補(as_owner):
    page = as_owner.get("/manage/news/new/").content.decode()
    for c in ["お知らせ", "ブログ", "休診情報", "イベント", "商品情報"]:
        assert f'value="{c}"' in page


def test_投稿すると行番号は最大プラス1で旧管理アプリからも見える(as_owner, news):
    r = as_owner.post("/manage/news/new/", _form())
    assert r.status_code == 302, r.content.decode()[:800]
    made = News.objects.get(title="マッサージ教室 おさらい会")
    assert made.sheet_row == 11 and made.published is True and made.notice_listed is True
    assert made.posted_on.isoformat() == "2026-09-14"  # 日時欄の時刻は落として日付だけ
    # GAS の getAdminBlogs と同じ一覧（旧管理アプリが読むもの）に載る
    rows = {b["rowIdx"]: b for b in 一覧()["blogs"]}
    assert rows[11]["title"] == "マッサージ教室 おさらい会" and rows[11]["status"] == "公開"
    page = as_owner.get("/manage/news/").content.decode()
    assert "お知らせを投稿しました！ 🎉" in page


def test_下書き保存は非公開でお知らせ一覧にも載らない(as_owner, categories):
    as_owner.post("/manage/news/new/", _form(status="非公開"))
    made = News.objects.get(title="マッサージ教室 おさらい会")
    assert made.published is False and made.notice_listed is False and made.notice_listed_at is None


def test_タイトルが無ければ弾く(as_owner, categories):
    r = as_owner.post("/manage/news/new/", _form(title=""))
    assert r.status_code == 200 and "タイトルが必要です" in r.content.decode()
    assert News.objects.count() == 0


def test_リンクと公開開始日時が入る(as_owner, categories):
    as_owner.post("/manage/news/new/", _form(linkUrl="https://example.com/", linkButtonText="予約はこちら", publishAt="2026-12-01T09:00"))
    made = News.objects.get(title="マッサージ教室 おさらい会")
    assert made.link_url == "https://example.com/" and made.button_text == "予約はこちら"
    assert timezone.localtime(made.publish_at).strftime("%Y-%m-%dT%H:%M") == "2026-12-01T09:00"


def test_画像を上げると縮めて公開URLになる(as_owner, categories, settings):
    r = as_owner.post("/manage/news/new/", {**_form(), "images": [_png(), _png()]})
    assert r.status_code == 302, r.content.decode()[:800]
    made = News.objects.get(title="マッサージ教室 おさらい会")
    urls = made.image_url.split("\n")
    assert len(urls) == 2
    assert all(u.startswith("https://api.example.com/media/news/") and u.endswith(".jpg") for u in urls)
    name = urls[0].rsplit("/", 1)[1]
    saved = settings.MEDIA_ROOT / "news" / name
    assert saved.is_file()
    from PIL import Image

    assert max(Image.open(saved).size) <= 1600
    # お客様向けの形（getAdminBlogs）でも2枚とも出る
    rows = {b["rowIdx"]: b for b in 一覧()["blogs"]}
    assert rows[made.sheet_row]["imageUrls"] == urls
    # 一覧のアイコン欄は1枚目の画像
    assert f'<img class="thumb" src="{urls[0]}"' in as_owner.get("/manage/news/").content.decode()


# ========== 修正 ==========


def test_修正で画像を外し新しいものを足せる(as_owner, news):
    r = as_owner.post(
        f"/manage/news/{news.sheet_row}/",
        {**_form(title="米味噌作り（改）"), "images": [_png()]},  # keep_image を送らない＝外す
    )
    assert r.status_code == 302, r.content.decode()[:800]
    news.refresh_from_db()
    assert news.title == "米味噌作り（改）"
    assert "drive.google.com" not in news.image_url
    assert news.image_url.startswith("https://api.example.com/media/news/")
    assert "ブログを更新しました！" in as_owner.get("/manage/news/").content.decode()


def test_修正で画像を残せる(as_owner, news):
    keep = "https://drive.google.com/thumbnail?id=abc"
    as_owner.post(f"/manage/news/{news.sheet_row}/", {**_form(), "keep_image": [keep]})
    news.refresh_from_db()
    assert news.image_url == keep


def test_修正画面にいまの値が入っている(as_owner, news):
    news.updated_at = timezone.make_aware(timezone.datetime(2026, 9, 10, 15, 45))
    news.save()
    page = as_owner.get(f"/manage/news/{news.sheet_row}/").content.decode()
    assert 'value="米味噌作り"' in page and "drive.google.com/thumbnail?id=abc" in page
    assert 'value="2026-09-10T15:45"' in page  # 日時は更新日時（旧アプリの editBlog と同じ）
    assert 'value="イベント情報" selected' in page
    assert "📝 修正して保存する" in page and "キャンセル" in page
    # 修正時は「お知らせ一覧に反映」はオフ
    assert 'name="notice_refresh" aria-label="NEWS保存時にお知らせ一覧へ反映" checked' not in page


def test_修正でタイトルを空にすると弾く(as_owner, news):
    r = as_owner.post(f"/manage/news/{news.sheet_row}/", _form(title=""))
    assert r.status_code == 200
    news.refresh_from_db()
    assert news.title == "米味噌作り"


def test_お知らせ一覧に反映はチェックがあり公開のときだけ(as_owner, news):
    news.notice_listed = False
    news.save()
    as_owner.post(f"/manage/news/{news.sheet_row}/", _form(notice_refresh=""))
    news.refresh_from_db()
    assert news.notice_listed is False  # 未チェックなら通常保存のみ
    as_owner.post(f"/manage/news/{news.sheet_row}/", _form(status="非公開"))
    news.refresh_from_db()
    assert news.notice_listed is False  # 下書きなら載せない
    as_owner.post(f"/manage/news/{news.sheet_row}/", _form())
    news.refresh_from_db()
    assert news.notice_listed is True and news.notice_listed_at is not None


# ========== 公開状況・削除 ==========


def test_公開状況のselectで公開と非公開を変える(as_owner, news):
    as_owner.post(f"/manage/news/{news.sheet_row}/status/", {"status": "非公開"})
    news.refresh_from_db()
    assert news.published is False
    r = as_owner.post(f"/manage/news/{news.sheet_row}/status/", {"status": "公開"}, follow=True)
    news.refresh_from_db()
    assert news.published is True and "公開設定を変更しました" in r.content.decode()


def test_削除は消さずに印を付ける(as_owner, news):
    r = as_owner.post(f"/manage/news/{news.sheet_row}/delete/", follow=True)
    news.refresh_from_db()
    assert news.deleted is True and news.deleted_at is not None
    assert News.objects.filter(pk=news.pk).exists()
    assert "ブログを削除しました" in r.content.decode()
    assert as_owner.get(f"/manage/news/{news.sheet_row}/").status_code == 404


def test_一括削除は印だけ(as_owner, categories):
    for t in ("a", "b", "c"):
        as_owner.post("/manage/news/new/", _form(title=t))
    rows = list(News.objects.order_by("sheet_row").values_list("sheet_row", flat=True))
    r = as_owner.post("/manage/news/bulk-delete/", {"rows": [str(rows[0]), str(rows[2])]}, follow=True)
    assert News.objects.filter(deleted=False).count() == 1 and News.objects.count() == 3
    assert "2件のブログ記事を削除しました" in r.content.decode()
    r = as_owner.post("/manage/news/bulk-delete/", {}, follow=True)
    assert "削除する項目を選択してください。" in r.content.decode()


def test_公開切替と削除はGETでは動かない(as_owner, news):
    assert as_owner.get(f"/manage/news/{news.sheet_row}/status/").status_code == 405
    assert as_owner.get(f"/manage/news/{news.sheet_row}/delete/").status_code == 405
    assert as_owner.get("/manage/news/bulk-delete/").status_code == 405
    news.refresh_from_db()
    assert news.published is True and news.deleted is False


# ========== 通知 ==========


def test_通知の文言(as_owner, categories, fake_onesignal):
    as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "📝 マッサージ教室 おさらい会"
    assert fake_onesignal[-1]["contents"]["ja"] == "NEWSが更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "news"


def test_通知は下書きと公開開始前とチェック無しでは送らない(as_owner, categories, fake_onesignal):
    as_owner.post("/manage/news/new/", _form(title="a"))
    as_owner.post("/manage/news/new/", _form(title="b", send_push="on", status="非公開"))
    as_owner.post("/manage/news/new/", _form(title="c", send_push="on", publishAt="2099-01-01T09:00"))
    assert fake_onesignal == []


# ========== 置き場の守り ==========


def test_お知らせ画像は置き場の外を読めない(client):
    assert client.get("/media/news/..%2F..%2Fsettings.py").status_code == 404
    assert client.get("/media/news/nothing.jpg").status_code == 404
