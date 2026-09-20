"""メニュー管理。**中身は旧管理アプリ（#page-menus メニュー管理）と同じ**
（入力項目の並び・下書き一覧と公開済み一覧・更新日時の列・つまんで並べ替え・お知らせ反映と通知のチェック）。
読み書きは gasapi/admin_menu（GAS の転送先）と同じ。
"""

import io

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile
from django.utils import timezone

from apps.content.models import Menu, PushNotice
from apps.gasapi.admin_menu import 一覧 as メニュー一覧

pytestmark = pytest.mark.django_db


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (800, 600), (120, 160, 90)).save(buf, format="PNG")
    return SimpleUploadedFile("p.png", buf.getvalue(), content_type="image/png")


def _menu(**extra):
    d = {"name": "産後ケア（訪問）", "category": "産後ケア", "description": "説明", "reservationStatus": "予約受付中",
         "publishAt": "", "notice_refresh": "on", "status": "公開"}
    d.update(extra)
    return d


def test_公開で追加すると旧管理アプリの一覧にも出て表示順が入る(as_owner):
    r = as_owner.post("/manage/menus/new/", _menu())
    assert r.status_code == 302, r.content.decode()[:500]
    m = Menu.objects.get(name="産後ケア（訪問）")
    assert m.sheet_row == 2 and m.published is True and m.booking_status == "予約受付中" and m.notice_listed is True
    assert (m.sort_key or 0) > 1_000_000_000_000  # GAS と同じく「今のミリ秒」。新しいものがホームで上に来る
    rows = {x["rowIdx"]: x for x in メニュー一覧()["menus"]}
    assert rows[2]["publishStatus"] == "公開" and rows[2]["reservationStatus"] == "予約受付中"
    page = as_owner.get("/manage/menus/").content.decode()
    assert "メニューを追加しました" in page


def test_下書き保存は非公開でお知らせ反映を外せば載らない(as_owner):
    as_owner.post("/manage/menus/new/", _menu(status="非公開", notice_refresh=""))
    m = Menu.objects.get()
    assert m.published is False and m.notice_listed is False


def test_カテゴリとメニュー名が無いと保存しない(as_owner):
    r = as_owner.post("/manage/menus/new/", _menu(category=""))
    assert r.status_code == 200 and "カテゴリとメニュー名を入力してください" in r.content.decode()
    assert Menu.objects.count() == 0


def test_画像を足して外せる(as_owner):
    as_owner.post("/manage/menus/new/", {**_menu(), "images": [_png(), _png()]})
    m = Menu.objects.get()
    assert len(m.image_urls) == 2 and all("/media/menus/" in u for u in m.image_urls)
    keep = m.image_urls[0]
    as_owner.post(f"/manage/menus/{m.sheet_row}/", {**_menu(name="改"), "keep_image": [keep]})
    m.refresh_from_db()
    assert m.image_urls == [keep] and m.name == "改"


def test_一覧は下書きと公開済みの2つで列と空のときの文言が旧管理アプリと同じ(as_owner):
    page = as_owner.get("/manage/menus/").content.decode()
    for s in ["メニュー管理", "🔄 更新", "📁 下書き保存一覧", "🍴 公開済みメニュー一覧",
              "<th>画像</th>", "<th>カテゴリ</th>", "<th>メニュー内容</th>", "<th>更新日時</th>", "<th>操作</th>",
              "下書きはありません", "公開済みのメニューはありません"]:
        assert s in page, s
    # 集計カード・検索・一括削除・一覧からの公開切替は旧管理アプリに無い
    # 「公開する」は左メニューの公式サイトの項目にもあるので、本文（main）だけで見る
    本文 = page[page.find("<main"):]
    for s in ["一括削除", "検索", "公開する", "非公開にする", "▲", "▼"]:
        assert s not in 本文, s


def test_一覧の行は名前と説明と更新日時でカテゴリ無しは未設定(as_owner):
    as_owner.post("/manage/menus/new/", _menu(description="<strong>太字</strong>と<u>下線</u>\n2行目<script>x</script>"))
    Menu.objects.filter(name="産後ケア（訪問）").update(category="")
    as_owner.post("/manage/menus/new/", _menu(name="説明なし", description="", status="非公開"))
    page = as_owner.get("/manage/menus/").content.decode()
    m = Menu.objects.get(name="産後ケア（訪問）")
    assert "<strong>太字</strong>と<u>下線</u><br>2行目&lt;script&gt;x&lt;/script&gt;" in page
    assert "概要説明は未入力です" in page and "未設定" in page
    # **画面は日本時間で出す。** 世界標準時のまま見比べると、
    # 日本の午前0時〜9時（＝世界標準時の前日15時以降）に日付が1日ずれて落ちる。
    assert timezone.localtime(m.updated_at).strftime("%Y/%m/%d") in page
    assert 'draggable="true"' in page and "☰" in page and "☰ をつまんで" in page


def test_一覧はお客様のホームと同じ順で並ぶ(as_owner):
    as_owner.post("/manage/menus/new/", _menu(name="先"))
    as_owner.post("/manage/menus/new/", _menu(name="後"))
    page = as_owner.get("/manage/menus/").content.decode()
    assert page.index("後") < page.index("先")  # 新しく足したものが上（表示順が今のミリ秒）


def test_つまんで動かした順を保存すると表示順に入る(as_owner):
    for n in ("A", "B", "C"):
        as_owner.post("/manage/menus/new/", _menu(name=n))
    a, b, c = (Menu.objects.get(name=n) for n in ("A", "B", "C"))
    r = as_owner.post("/manage/menus/order/", {"rows": [str(c.sheet_row), str(a.sheet_row), str(b.sheet_row)]})
    assert r.status_code == 302
    a.refresh_from_db(); b.refresh_from_db(); c.refresh_from_db()
    # 旧管理アプリの saveMenuOrder と同じ計算: 上から (件数 - 位置) * 1000 + 1000000
    assert (c.sort_key, a.sort_key, b.sort_key) == (1003000, 1002000, 1001000)
    assert (a.sheet_row, b.sheet_row, c.sheet_row) == (2, 3, 4)  # 行番号は動かない
    page = as_owner.get("/manage/menus/").content.decode()
    assert "メニュー順序を保存しました" in page
    assert page.index('data-row-idx="4"') < page.index('data-row-idx="2"') < page.index('data-row-idx="3"')


def test_削除は印だけで一覧から消える(as_owner):
    as_owner.post("/manage/menus/new/", _menu(name="消すもの"))
    m = Menu.objects.get()
    as_owner.post(f"/manage/menus/{m.sheet_row}/delete/")
    m.refresh_from_db()
    assert m.deleted is True and Menu.objects.filter(pk=m.pk).exists()
    page = as_owner.get("/manage/menus/").content.decode()
    assert "「消すもの」を削除しました" in page and 'data-row-idx=' not in page


def test_編集でお知らせ反映を外すと掲載は変わらずチェックすると載せ直す(as_owner):
    as_owner.post("/manage/menus/new/", _menu())
    m = Menu.objects.get()
    assert m.notice_listed is True
    Menu.objects.filter(pk=m.pk).update(notice_listed=False)
    as_owner.post(f"/manage/menus/{m.sheet_row}/", _menu(name="修正", notice_refresh=""))
    m.refresh_from_db()
    assert m.name == "修正" and m.notice_listed is False  # 未チェックならメニュー情報だけ更新
    as_owner.post(f"/manage/menus/{m.sheet_row}/", _menu(name="修正2", notice_refresh="on"))
    m.refresh_from_db()
    assert m.notice_listed is True and m.notice_listed_at is not None
    page = as_owner.get("/manage/menus/").content.decode()
    assert "メニューを更新しました" in page


def test_入力画面の項目と文言(as_owner):
    page = as_owner.get("/manage/menus/new/").content.decode()
    for s in ["✨ 新しいメニューを追加する", "カテゴリ", 'placeholder="例：産後ケア"', "メニュー名", 'placeholder="例：骨盤ケア・整体"',
              "概要説明", "選択した文字に太字・下線を設定できます", "画像（フライヤー・写真など）をアップロード",
              "※画像全体を保ったまま、アップロード用に最適化して取り込みます", "予約受付ステータス", "予約受付中", "予約対象外",
              "公開開始日時（空欄ならすぐ公開）", "この保存をお知らせ一覧に反映する",
              "初回掲載はオンのまま、通常の説明修正時は必要なときだけチェックしてください。",
              "新規追加は既定でオンです。", "この保存でPush通知を送信する", "未チェックでは通知されません。",
              "公開保存のときだけ送信対象になります。", "👁 プレビュー", "🍴 公開して保存する", "📁 下書き保存", "メニュープレビュー"]:
        assert s in page, s
    assert 'name="notice_refresh" checked' in page and "キャンセル" not in page
    # 項目の並びが旧管理アプリと同じ（カテゴリ → メニュー名 → 概要説明 → 画像 → 予約 → 公開開始 → お知らせ → Push）
    順 = [page.index(s) for s in ['id="m-cat"', 'id="m-name"', 'id="m-desc"', 'id="m-img"', 'id="m-res"', 'id="m-at"', 'id="m-notice"', 'id="m-push"']]
    assert 順 == sorted(順)

    as_owner.post("/manage/menus/new/", _menu())
    m = Menu.objects.get()
    page = as_owner.get(f"/manage/menus/{m.sheet_row}/").content.decode()
    assert "📝 メニューを編集する" in page and "🍴 変更を保存する" in page and "キャンセル" in page
    assert 'name="notice_refresh" checked' not in page  # 編集はオフから


def test_通知は公開時にチェックしたときだけ(as_owner, fake_onesignal):
    as_owner.post("/manage/menus/new/", _menu(name="通知あり", send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "🍴 通知あり"
    assert fake_onesignal[-1]["contents"]["ja"] == "ホームのメニュー一覧が更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "home"
    n = len(fake_onesignal)
    as_owner.post("/manage/menus/new/", _menu(name="下書き", status="非公開", send_push="on"))
    assert len(fake_onesignal) == n
    assert PushNotice.objects.count() == 1


def test_一覧からの公開切替と上下ボタンは旧管理アプリに無いので無い(as_owner):
    as_owner.post("/manage/menus/new/", _menu())
    m = Menu.objects.get()
    assert as_owner.post(f"/manage/menus/{m.sheet_row}/toggle/").status_code == 404
    assert as_owner.post(f"/manage/menus/{m.sheet_row}/move/", {"direction": "up"}).status_code == 404


def test_プレビューは1か所失敗しても止まらない(as_owner):
    """院長の画面で、札も名前も説明も空のまま画像だけ壊れて出ていた（2026-09-20）。

    1か所でも失敗すると後ろが全部止まる作りだったので、欄ごとに分けた。
    """
    as_owner.post("/manage/menus/new/", _menu())
    m = Menu.objects.get(name="産後ケア（訪問）")
    page = as_owner.get(f"/manage/menus/{m.sheet_row}/").content.decode()
    assert "欄に入れる_" in page          # 欄ごとに分けて入れる
    assert "img.onerror" in page          # 出せない画像は出さない
    assert "うまく出せませんでした" in page  # 失敗したら理由を出す
