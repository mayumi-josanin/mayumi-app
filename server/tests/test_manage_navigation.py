"""左メニューのまとまりと上部タブ（院長の相談 2026-09-16: 入口を減らし、中はタブで切り替える）。"""

import pytest

pytestmark = pytest.mark.django_db


def _sidebar(page):
    return page[page.find('<nav class="sidebar-nav">'):page.find("</nav>")]


def _tabs(page):
    i = page.find('<nav class="page-tabs"')
    return page[i:page.find("</nav>", i)] if i != -1 else ""


def test_左メニューは12のまとまり(as_owner):
    page = as_owner.get("/manage/orders/").content.decode()
    side = _sidebar(page)
    for g in ["会員", "アプリの掲載", "通知", "注文と売上", "設定", "記事を書く", "ページを整える", "確かめて公開"]:
        assert f" {g}</a>" in side, g
    # 昔の個別の入口は左メニューには無い（中のタブに移った）
    for old in [" NEWS</a>", " 商品管理</a>", " カレンダー</a>", " スタンプ・特典</a>", " ゴミ箱</a>", " お知らせ管理</a>"]:
        assert old not in side, old
    assert ">アプリ管理</span>" in side and ">公式サイト</span>" in side and ">予約管理</span>" in side


def test_開いている画面のまとまりが光り上部タブが出る(as_owner):
    page = as_owner.get("/manage/members/").content.decode()
    side = _sidebar(page)
    assert 'class="active"><span class="nav-icon">👥</span> 会員</a>' in side
    tabs = _tabs(page)
    assert 'class="active">会員一覧</a>' in tabs and "スタンプ・特典" in tabs and "重複候補" in tabs
    page = as_owner.get("/manage/rewards/").content.decode()
    assert 'class="active">スタンプ・特典</a>' in _tabs(page)
    page = as_owner.get("/manage/news/").content.decode()
    assert 'class="active"><span class="nav-icon">📱</span> アプリの掲載</a>' in _sidebar(page)
    assert 'class="active">NEWS（アプリ）</a>' in _tabs(page) and "お知らせ一覧の見え方" in _tabs(page)
    page = as_owner.get("/manage/revenue/menu/").content.decode()
    assert 'class="active">売上の記録（メニュー）</a>' in _tabs(page)
    assert 'class="active"><span class="nav-icon">📦</span> 注文と売上</a>' in _sidebar(page)
    page = as_owner.get("/manage/system/").content.decode()
    assert 'class="active">システム管理</a>' in _tabs(page)


def test_公式サイトのまとまり(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    page = as_owner.get("/manage/hp/publish/").content.decode()
    assert 'class="active"><span class="nav-icon">🌐</span> 確かめて公開</a>' in _sidebar(page)
    assert 'class="active">公開する</a>' in _tabs(page) and "保存の記録" in _tabs(page)
    page = as_owner.get("/manage/hp/news/").content.decode()
    assert 'class="active">お知らせ（サイト）</a>' in _tabs(page)
