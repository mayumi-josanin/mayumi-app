"""お知らせ管理。旧管理アプリ（GAS 経由）と同じ表を、直接読み書きする。"""

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
        "posted_on": "2026-09-14", "title": "マッサージ教室 おさらい会", "category": "イベント情報",
        "icon": "📢", "body": "9月14日 10:00〜11:30", "publish_at": "", "link_url": "",
        "button_text": "", "notice_listed": "on", "status": "公開",
    }
    data.update(extra)
    return data


# ========== 一覧 ==========


def test_一覧に投稿済みと下書きが分かれて出る(as_owner, news):
    News.objects.create(sheet_row=11, posted_on=timezone.localdate(), title="下書きの記事",
                        category="まゆみのつぶやき", published=False)
    News.objects.create(sheet_row=12, posted_on=timezone.localdate(), title="消した記事",
                        category="イベント情報", published=True, deleted=True)
    page = as_owner.get("/manage/news/").content.decode()
    assert "米味噌作り" in page
    assert "下書きの記事" in page
    assert "消した記事" not in page
    # 種別はカテゴリ側が持つ
    assert "badge-red" in page and "badge-green" in page


def test_種別で絞れる(as_owner, news):
    News.objects.create(sheet_row=11, posted_on=timezone.localdate(), title="つぶやき",
                        category="まゆみのつぶやき", published=True)
    page = as_owner.get("/manage/news/?kind=ブログ").content.decode()
    assert "つぶやき" in page
    assert "米味噌作り" not in page


# ========== 投稿 ==========


def test_投稿すると行番号は最大プラス1で旧管理アプリからも見える(as_owner, news):
    r = as_owner.post("/manage/news/new/", _form())
    assert r.status_code == 302, r.content.decode()[:800]
    made = News.objects.get(title="マッサージ教室 おさらい会")
    assert made.sheet_row == 11
    assert made.published is True
    assert made.notice_listed is True
    # GAS の getAdminBlogs と同じ一覧（旧管理アプリが読むもの）に載る
    rows = {b["rowIdx"]: b for b in 一覧()["blogs"]}
    assert rows[11]["title"] == "マッサージ教室 おさらい会"
    assert rows[11]["status"] == "公開"


def test_下書き保存は非公開(as_owner, categories):
    as_owner.post("/manage/news/new/", _form(status="非公開"))
    assert News.objects.get(title="マッサージ教室 おさらい会").published is False


def test_タイトルが無ければ弾く(as_owner, categories):
    r = as_owner.post("/manage/news/new/", _form(title=""))
    assert r.status_code == 200
    assert News.objects.count() == 0


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


def test_修正で画像を残せる(as_owner, news):
    keep = "https://drive.google.com/thumbnail?id=abc"
    as_owner.post(f"/manage/news/{news.sheet_row}/", {**_form(), "keep_image": [keep]})
    news.refresh_from_db()
    assert news.image_url == keep


def test_修正画面にいまの値が入っている(as_owner, news):
    page = as_owner.get(f"/manage/news/{news.sheet_row}/").content.decode()
    assert 'value="米味噌作り"' in page
    assert "drive.google.com/thumbnail?id=abc" in page


# ========== 公開切替・削除 ==========


def test_公開と非公開を切り替える(as_owner, news):
    as_owner.post(f"/manage/news/{news.sheet_row}/toggle/")
    news.refresh_from_db()
    assert news.published is False
    as_owner.post(f"/manage/news/{news.sheet_row}/toggle/")
    news.refresh_from_db()
    assert news.published is True


def test_削除は消さずに印を付ける(as_owner, news):
    as_owner.post(f"/manage/news/{news.sheet_row}/delete/")
    news.refresh_from_db()
    assert news.deleted is True
    assert news.deleted_at is not None
    assert News.objects.filter(pk=news.pk).exists()
    assert as_owner.get(f"/manage/news/{news.sheet_row}/").status_code == 404


def test_切替と削除はGETでは動かない(as_owner, news):
    assert as_owner.get(f"/manage/news/{news.sheet_row}/toggle/").status_code == 405
    assert as_owner.get(f"/manage/news/{news.sheet_row}/delete/").status_code == 405
    news.refresh_from_db()
    assert news.published is True and news.deleted is False


# ========== 置き場の守り ==========


def test_お知らせ画像は置き場の外を読めない(client):
    assert client.get("/media/news/..%2F..%2Fsettings.py").status_code == 404
    assert client.get("/media/news/nothing.jpg").status_code == 404
