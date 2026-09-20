"""カレンダー管理。**中身は旧管理アプリ（カレンダー管理）と同じ**
（日付の複数追加・年/月の絞り込み・下書きとイベントの二段の一覧・月間カレンダー・
最近使った色・対象メニューはイベントだけ・公開設定・一括削除・通知）。
読み書きは gasapi/admin_calendar（GAS の転送先）と同じ。"""

import io
from datetime import date

import pytest

from apps.content.models import CalendarEvent, Menu, PushNotice
from apps.gasapi.admin_calendar import 一覧

pytestmark = pytest.mark.django_db


def _form(**extra):
    d = {"date": "2026-10-03", "title": "ベビーマッサージ教室", "category": "教室", "menuRowIdx": "",
         "desc": "10:00〜", "color": "#f48fb1", "publishAt": "", "linkUrl": "", "linkButtonText": "",
         "notice_refresh": "on", "status": "公開"}
    d.update(extra)
    return d


def test_日付を複数選ぶと選んだ数だけ作る(as_owner):
    r = as_owner.post("/manage/calendar/new/", {**_form(), "dates": ["2026-10-03", "2026-10-10", "2026-10-17"]})
    assert r.status_code == 302, r.content.decode()[:500]
    rows = list(CalendarEvent.objects.order_by("sheet_row"))
    assert [c.sheet_row for c in rows] == [2, 3, 4]
    assert [c.event_on for c in rows] == [date(2026, 10, 3), date(2026, 10, 10), date(2026, 10, 17)]
    assert all(c.title == "ベビーマッサージ教室" and c.color == "#f48fb1" and c.published and c.notice_listed for c in rows)
    assert len(一覧()["events"]) == 3


def test_追加した日が無ければ日付1つで作る(as_owner):
    as_owner.post("/manage/calendar/new/", _form())
    assert CalendarEvent.objects.count() == 1


def test_日付が無ければ断り足した日は残す(as_owner):
    r = as_owner.post("/manage/calendar/new/", _form(date=""))
    assert r.status_code == 200 and "日付を少なくとも1つ選択（追加）してください。" in r.content.decode()
    assert CalendarEvent.objects.count() == 0
    # イベント名が空なら gasapi が断る。足してあった日付は画面に残す
    r = as_owner.post("/manage/calendar/new/", {**_form(title=""), "dates": ["2026-10-03"]})
    assert r.status_code == 200 and '<script id="picked-dates" type="application/json">["2026-10-03"]</script>' in r.content.decode()


def test_追加画面の項目と文言(as_owner):
    page = as_owner.get("/manage/calendar/new/").content.decode()
    for s in ["✨ 新しいイベントを追加する", "日時選択", "選択中の日付", "イベント名", "例：産後ヨガ教室",
              "例：教室 / 休診 / イベント", "なし（休診・往診など）",
              "選ぶと、お客様アプリのメニュー一覧に「◯月◯日開催予定」と表示されます",
              "選択した文字に太字・下線を設定できます", "10:00〜11:30 定員5名様など...",
              "★画像（フライヤー・写真など）をアップロード", "カラー設定",
              "例：休診日(赤系 #e57373)、ヨガ(ピンク系 #f48fb1)", "使用履歴のある色がここに表示されます",
              "公開開始日時（空欄ならすぐ公開）", "リンクURL（任意：イベント詳細ボタンの遷移先）",
              "ボタンのテキスト（任意）", "詳しく見る （未入力ならこれになります）",
              "この保存をお知らせ一覧に反映する", "新規追加は既定でオンです。",
              "この保存でPush通知を送信する", "未チェックでは通知されません。",
              "👁 プレビュー", "📅 公開して追加する", "📁 下書き保存"]:
        assert s in page, s
    assert 'name="notice_refresh" checked' in page  # 新規は既定でオン
    assert "datalist" not in page  # 旧アプリのカテゴリは自由入力


def test_最近使った色は全イベントから重複なく(as_owner):
    as_owner.post("/manage/calendar/new/", _form(color="#F48FB1"))
    as_owner.post("/manage/calendar/new/", _form(title="休診日", color="#e57373", date="2026-10-04"))
    as_owner.post("/manage/calendar/new/", _form(title="もう一つ", color="#f48fb1", date="2026-10-05"))
    page = as_owner.get("/manage/calendar/new/").content.decode()
    assert "使用履歴から選択:" in page
    assert page.count('class="color-swatch"') == 2
    assert "setCalendarFormColor('#f48fb1')" in page and "setCalendarFormColor('#e57373')" in page


def test_対象メニューはイベントだけ_結びついているものは残す(as_owner):
    Menu.objects.create(sheet_row=5, name="ベビマ", category="イベント", published=True)
    Menu.objects.create(sheet_row=6, name="母乳相談", category="通常", published=True)
    Menu.objects.create(sheet_row=7, name="よもぎ蒸し", category="通常", published=True)
    page = as_owner.get("/manage/calendar/new/").content.decode()
    assert "ベビマ" in page and "母乳相談" not in page and "よもぎ蒸し" not in page
    as_owner.post("/manage/calendar/new/", _form(menuRowIdx="6"))
    c = CalendarEvent.objects.get()
    assert c.menu_row == 6
    page = as_owner.get(f"/manage/calendar/{c.sheet_row}/").content.decode()
    assert "📝 イベントを編集する" in page and "📝 修正を保存する" in page
    assert 'value="6" selected' in page and "母乳相談" in page and "よもぎ蒸し" not in page
    assert 'name="notice_refresh" checked' not in page  # 編集は既定でオフ
    assert {e["rowIdx"]: e["menuRowIdx"] for e in 一覧()["events"]}[c.sheet_row] == 6


def test_一覧の列と並びと空の文言(as_owner):
    page = as_owner.get("/manage/calendar/").content.decode()
    for s in ["カレンダー管理", "🗑 一括削除", "🔄 更新", "下書き保存一覧", "イベント一覧", "📅 月別フィルタ",
              "全年度", "全月", "条件をクリア", "下書き保存はありません", "投稿済みイベントはありません",
              "表示条件に一致するイベントはありません", "凡例:", "訪問産後ケア", "※日付を押すと、その日の予定と「この日に追加する」が出ます",
              "← 前月", "今月", "翌月 →"]:
        assert s in page, s
    for th in ["<th>画像</th>", "<th>イベント名</th>", "<th>カテゴリ</th>", "<th>詳細</th>", "<th>カラー</th>", "<th>公開設定</th>", "<th>操作</th>"]:
        assert th in page, th
    as_owner.post("/manage/calendar/new/", _form(title="早い日", date="2026-10-03", category="", desc=""))
    as_owner.post("/manage/calendar/new/", _form(title="遅い日", date="2026-10-20", desc="<strong>太字</strong>の説明 " + "あ" * 80))
    as_owner.post("/manage/calendar/new/", _form(title="下書き", date="2026-10-10", status="非公開"))
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert page.index("遅い日") < page.index("早い日")  # 日付の降順
    assert "未設定" in page and "―" in page  # カテゴリ空・詳細空
    assert "太字の説明 " + "あ" * 54 + "…" in page and "<strong>太字</strong>の説明" not in page
    assert "3 / 3 件を表示" in page
    assert page.index("下書き") < page.index("遅い日")  # 下書き保存一覧が上


def test_一覧は年月で絞れる_既定は今月(as_owner):
    as_owner.post("/manage/calendar/new/", _form(date="2026-10-03"))
    as_owner.post("/manage/calendar/new/", _form(title="11月の予定", date="2026-11-03"))
    as_owner.post("/manage/calendar/new/", _form(title="来年の予定", date="2027-01-03"))
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert "ベビーマッサージ教室" in page and "11月の予定" not in page and "1 / 3 件を表示" in page
    page = as_owner.get("/manage/calendar/?year=2026&month=all").content.decode()
    assert "11月の予定" in page and "来年の予定" not in page
    page = as_owner.get("/manage/calendar/?year=all&month=all").content.decode()
    assert "来年の予定" in page and "3 / 3 件を表示" in page
    # 既定は今年・今月（院長の希望 2026-09-16: 一覧は月ごとに見る）
    from django.utils import timezone

    today = timezone.localdate()
    page = as_owner.get("/manage/calendar/").content.decode()
    assert f'<option value="{today.year}" selected>' in page and f'<option value="{today.month}" selected>{today.month}月' in page
    assert f"{today.year + 1}年" in page and f"{today.year - 1}年" in page


def test_月間カレンダーには公開中だけ渡す(as_owner):
    as_owner.post("/manage/calendar/new/", _form(title="休診日", category="休診", date="2026-10-04"))
    as_owner.post("/manage/calendar/new/", _form(title="下書き", date="2026-10-10", status="非公開"))
    as_owner.post("/manage/calendar/new/", _form(title="往診の日", category="", date="2026-10-11"))
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert 'id="cal-events"' in page and "monthCalendarState = { year: 2026, month: 10 }" in page
    import json
    import re

    data = json.loads(re.search(r'<script id="cal-events" type="application/json">(.*?)</script>', page).group(1))
    by = {e["title"]: e for e in data}
    assert by["休診日"]["kind"] == "holiday" and by["往診の日"]["kind"] == "visit" and by["下書き"]["status"] == "非公開"
    assert by["休診日"]["editUrl"].endswith(f"/manage/calendar/{by['休診日']['rowIdx']}/")


def test_修正は日付1つだけを直しお知らせ反映は印のときだけ(as_owner):
    as_owner.post("/manage/calendar/new/", {**_form(), "dates": ["2026-10-03", "2026-10-10"]})
    c = CalendarEvent.objects.get(event_on=date(2026, 10, 10))
    CalendarEvent.objects.filter(pk=c.pk).update(notice_listed=False)
    r = as_owner.post(f"/manage/calendar/{c.sheet_row}/", _form(date="2026-10-11", title="改", dates=["2026-12-01"], notice_refresh=""))
    assert r.status_code == 302 and r["Location"] == "/manage/calendar/?year=2026&month=10"
    c.refresh_from_db()
    assert c.event_on == date(2026, 10, 11) and c.title == "改"
    assert c.notice_listed is False and c.notice_listed_at is None  # 未チェックなら掲載は触らない
    assert CalendarEvent.objects.count() == 2  # dates は修正では無視する
    as_owner.post(f"/manage/calendar/{c.sheet_row}/", _form(date="2026-10-11", title="改", notice_refresh="on"))
    c.refresh_from_db()
    assert c.notice_listed is True and c.notice_listed_at is not None


def test_公開設定の切替と削除は印だけ(as_owner):
    as_owner.post("/manage/calendar/new/", _form())
    c = CalendarEvent.objects.get()
    r = as_owner.post(f"/manage/calendar/{c.sheet_row}/status/", {"status": "非公開", "next": "/manage/calendar/?year=2026&month=10"})
    assert r["Location"] == "/manage/calendar/?year=2026&month=10"
    c.refresh_from_db()
    assert c.published is False
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert "公開設定を変更しました" in page and '<option value="非公開" selected>' in page
    as_owner.post(f"/manage/calendar/{c.sheet_row}/delete/")
    c.refresh_from_db()
    assert c.deleted is True and CalendarEvent.objects.filter(pk=c.pk).exists()
    assert 一覧()["events"] == []


def test_一括削除は印だけ(as_owner):
    for d in ("2026-10-01", "2026-10-02", "2026-10-03"):
        as_owner.post("/manage/calendar/new/", _form(date=d))
    rows = list(CalendarEvent.objects.order_by("sheet_row").values_list("sheet_row", flat=True))
    r = as_owner.post("/manage/calendar/bulk-delete/", {"rows": [str(rows[0]), str(rows[2])]})
    assert r.status_code == 302
    assert CalendarEvent.objects.filter(deleted=False).count() == 1 and CalendarEvent.objects.count() == 3
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert "2件のイベントを削除しました" in page
    r = as_owner.post("/manage/calendar/bulk-delete/", {}, follow=True)
    assert "削除する項目を選択してください。" in r.content.decode()


def test_通知の文言と条件(as_owner, fake_onesignal):
    as_owner.post("/manage/calendar/new/", _form(send_push="on"))
    assert fake_onesignal[-1]["headings"]["ja"] == "📅 ベビーマッサージ教室"
    assert fake_onesignal[-1]["contents"]["ja"] == "カレンダーが更新されました"
    assert fake_onesignal[-1]["data"]["openPage"] == "calendar"
    assert PushNotice.objects.count() == 1
    as_owner.post("/manage/calendar/new/", _form(title="下書き", status="非公開", send_push="on"))
    assert len(fake_onesignal) == 1
    as_owner.post("/manage/calendar/new/", _form(title="未チェック"))
    assert len(fake_onesignal) == 1


def test_画像はカレンダーの置き場へ(as_owner):
    from django.core.files.uploadedfile import SimpleUploadedFile
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (400, 300), (10, 20, 30)).save(buf, format="PNG")
    as_owner.post("/manage/calendar/new/", {**_form(), "images": [SimpleUploadedFile("a.png", buf.getvalue(), content_type="image/png")]})
    c = CalendarEvent.objects.get()
    assert "/media/calendar/" in c.image_url
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    assert 'class="thumb"' in page
    assert as_owner.get("/media/calendar/nothing.jpg").status_code == 404


def test_月の表から日を押すとその日で追加できる(as_owner):
    """一覧の月の表 → 「この日に追加する」 → その日が入った追加の画面（院長の依頼 2026-09-20）。"""
    page = as_owner.get("/manage/calendar/").content.decode()
    # 予定の有無にかかわらず、どの日も押せて詳細が開く（空の日こそ、その日で追加したい）
    assert 'id="dayDetailAdd"' in page and "➕ この日に追加する" in page
    assert "CAL_NEW_URL = '/manage/calendar/new/'" in page
    assert "onDay: openDayDetailModal" in page
    # 押した日を ?date= に載せている
    assert "CAL_NEW_URL + '?date=' + encodeURIComponent(dateStr)" in page


def test_予定の無い日でも追加の入口を出す(as_owner):
    as_owner.post("/manage/calendar/new/", _form(date="2026-10-03"))
    page = as_owner.get("/manage/calendar/?year=2026&month=10").content.decode()
    # 押せるかどうかを予定の有無で分けていない（renderMonthCalendar に onDay を必ず渡す）
    assert "onDay: openDayDetailModal" in page
    assert "この日にイベントはありません" in page and "➕ この日に追加する" in page


def test_追加の画面は日付の指定を受け取る(as_owner):
    page = as_owner.get("/manage/calendar/new/?date=2026-11-23").content.decode()
    assert 'id="c-date" name="date" type="date" class="form-control" value="2026-11-23"' in page
    assert '<script id="picked-dates" type="application/json">["2026-11-23"]</script>' in page
    # 指定が無ければ今日を初期値にするが、選んだ日は空のまま
    page = as_owner.get("/manage/calendar/new/").content.decode()
    assert '<script id="picked-dates" type="application/json">[]</script>' in page
    # おかしな指定は今日に戻す（落とさない）
    assert as_owner.get("/manage/calendar/new/?date=2026-13-40").status_code == 200
    assert as_owner.get("/manage/calendar/new/?date=あああ").status_code == 200


def test_追加の画面に月の表が出る(as_owner):
    as_owner.post("/manage/calendar/new/", _form(title="先にある予定", date="2026-10-03"))
    as_owner.post("/manage/calendar/new/", _form(title="下書き", date="2026-10-05", status="非公開"))
    page = as_owner.get("/manage/calendar/new/?date=2026-10-03").content.decode()
    assert 'id="pickerGrid"' in page and "month_calendar.js" in page
    assert "日時選択（カレンダーの日付を押して選びます。いくつでも選べます）" in page
    assert "← 前月" in page and "翌月 →" in page
    assert "同じ曜日をまとめて選ぶ:" in page and "毎週月" in page
    assert "カレンダーの日付を押して選んでください" in page
    import json
    import re

    data = json.loads(re.search(r'<script id="picker-events" type="application/json">(.*?)</script>', page).group(1))
    # 月の表に出すのは公開中だけ（一覧の月の表と同じ考え方）
    assert [e["title"] for e in data] == ["先にある予定"]


def test_月をまたいで選んだ日の数だけ行を作る(as_owner):
    r = as_owner.post("/manage/calendar/new/", {**_form(date="2026-10-03"),
                                                "dates": ["2026-10-03", "2026-11-07", "2027-01-09"]})
    assert r.status_code == 302
    assert [c.event_on for c in CalendarEvent.objects.order_by("sheet_row")] == [
        date(2026, 10, 3), date(2026, 11, 7), date(2027, 1, 9)]


def test_修正のときは複数選択を出さない(as_owner):
    as_owner.post("/manage/calendar/new/", _form())
    c = CalendarEvent.objects.get()
    page = as_owner.get(f"/manage/calendar/{c.sheet_row}/").content.decode()
    assert 'id="pickerGrid"' not in page and "選択中の日付" not in page
    assert "同じ曜日をまとめて選ぶ:" not in page
    assert 'id="c-date" name="date" type="date" class="form-control" required' in page
