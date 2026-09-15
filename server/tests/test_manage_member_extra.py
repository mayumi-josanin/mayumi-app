"""重複会員候補（統合）と、バックアップ / ゴミ箱。
**中身は旧管理アプリ（会員管理ページの下半分: 🧩 重複会員候補・🛟 バックアップ / ゴミ箱）と同じ**
（判定理由3種・新しい会員へ統合・安全な候補の一括統合・ゴミ箱の検索と種類・復元・完全削除・一括）。"""

from datetime import date, timedelta

import pytest
from django.utils import timezone

from apps.content.models import CalendarEvent, Menu, News, Product, PushNotice
from apps.gasapi import trash
from apps.manage import member_gate, views_member_extra
from apps.members.models import Member
from apps.records.models import OrderLine

pytestmark = pytest.mark.django_db


def _member(mid, name, phone="", birthday=None, days_ago=0, **extra):
    return Member.objects.create(member_id=mid, name=name, phone=phone, birthday=birthday,
                                 created_at=timezone.now() - timedelta(days=days_ago), **extra)


# ---- 重複会員候補 ----

def test_まゆみ以外は入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/members/duplicates/").status_code == 403
    assert client.get("/manage/trash/").status_code == 403


def test_判定理由は電話番号と氏名生年月日と氏名の3種で同じ顔ぶれは1組(as_owner):
    # 電話も氏名も生年月日も同じ2人 → 1組に理由が3つ並ぶ
    _member("MYM-0001", "山田 花子", phone="090-1111-2222", birthday=date(1990, 1, 2), days_ago=10)
    _member("MYM-0002", "山田花子", phone="09011112222", birthday=date(1990, 1, 2), days_ago=1)
    # 氏名だけ同じ
    _member("MYM-0003", "佐藤 桃子", days_ago=5)
    _member("MYM-0004", "佐藤桃子", days_ago=2)
    # 退会した方は数えない
    _member("MYM-0005", "佐藤桃子", deleted=True)
    _member("MYM-0006", "ひとり", phone="080-0000-0000")
    page = as_owner.get("/manage/members/duplicates/").content.decode()
    for label in ["🧩 重複会員候補", "判定理由", "候補会員", "統合", "安全な候補を一括統合", "🔄 更新", "比較する", "最新の会員IDへ統合"]:
        assert label in page, label
    assert "電話番号が一致 / 氏名と生年月日が一致 / 氏名が一致" in page
    assert "氏名が一致</td>" in page
    assert "候補 2組 / 一括統合できる安全な候補 1組" in page
    assert "MYM-0005" not in page and "ひとり" not in page
    # 組の中はいちばん新しい会員が先頭（統合先）
    組 = views_member_extra.重複候補([{"memberId": "MYM-0001", "name": "山田花子", "phone": "09011112222", "timestamp": "2026-01-01T00:00:00+09:00"},
                                    {"memberId": "MYM-0002", "name": "山田花子", "phone": "09011112222", "timestamp": "2026-02-01T00:00:00+09:00"}])
    assert 組[0]["target"]["memberId"] == "MYM-0002" and 組[0]["sources"] == ["MYM-0001"]


def test_候補が無いときの文言(as_owner):
    page = as_owner.get("/manage/members/duplicates/").content.decode()
    assert "重複候補はありません" in page
    assert "同じ電話番号、氏名と生年月日、氏名が一致する候補を一覧表示します" in page


def test_切り替え前は統合を断る(as_owner):
    _member("MYM-0001", "山田花子", phone="09011112222", days_ago=3)
    _member("MYM-0002", "山田花子", phone="09011112222")
    r = as_owner.post("/manage/members/duplicates/merge/", {"targetMemberId": "MYM-0002", "sourceMemberIds": ["MYM-0001"]}, follow=True)
    assert member_gate.断る文() in r.content.decode()
    assert Member.objects.get(pk="MYM-0001").deleted is False
    page = as_owner.get("/manage/members/duplicates/").content.decode()
    assert "旧管理アプリで行ってください" in page


def test_統合すると元の会員はゴミ箱へ移り注文も付け替わる(as_owner):
    member_gate.切り替える("server")
    _member("MYM-0001", "山田花子", phone="09011112222", days_ago=3, memo="古いメモ", stamp_count=4)
    _member("MYM-0002", "山田花子", days_ago=0, stamp_count=2)
    OrderLine.objects.create(order_id="ORD-1", product_name="茶", quantity=1, member_id="MYM-0001")
    r = as_owner.post("/manage/members/duplicates/merge/", {"targetMemberId": "MYM-0002", "sourceMemberIds": ["MYM-0001"]}, follow=True)
    assert "会員を統合しました" in r.content.decode()
    先, 元 = Member.objects.get(pk="MYM-0002"), Member.objects.get(pk="MYM-0001")
    assert 元.deleted and 元.merged_into_id == "MYM-0002"
    assert 先.phone == "09011112222" and 先.memo == "古いメモ" and 先.stamp_count == 4
    assert OrderLine.objects.get(order_id="ORD-1").member_id == "MYM-0002"
    # ゴミ箱に「会員」として並び、統合先が理由に出る
    page = as_owner.get("/manage/trash/").content.decode()
    assert "MYM-0002 へ統合" in page and ">会員<" in page
    # 候補からは消える
    assert "重複候補はありません" in as_owner.get("/manage/members/duplicates/").content.decode()


def test_安全な候補の一括統合は氏名だけ一致を残す(as_owner):
    member_gate.切り替える("server")
    _member("MYM-0001", "山田花子", phone="09011112222", days_ago=3)
    _member("MYM-0002", "山田 花子", phone="090-1111-2222", days_ago=1)
    _member("MYM-0003", "鈴木一郎", birthday=date(1985, 5, 5), days_ago=9)
    _member("MYM-0004", "鈴木 一郎", birthday=date(1985, 5, 5), days_ago=4)
    _member("MYM-0005", "佐藤桃子", days_ago=5)
    _member("MYM-0006", "佐藤 桃子", days_ago=2)
    r = as_owner.post("/manage/members/duplicates/merge/", {"mode": "safe"}, follow=True)
    assert "安全な重複候補を 2組 (2件) 統合しました" in r.content.decode()
    assert Member.objects.get(pk="MYM-0001").merged_into_id == "MYM-0002"
    assert Member.objects.get(pk="MYM-0003").merged_into_id == "MYM-0004"
    assert Member.objects.get(pk="MYM-0005").deleted is False and Member.objects.get(pk="MYM-0006").deleted is False
    page = as_owner.get("/manage/members/duplicates/").content.decode()
    assert "氏名が一致" in page and "候補 1組 / 一括統合できる安全な候補 0組" in page
    r = as_owner.post("/manage/members/duplicates/merge/", {"mode": "safe"}, follow=True)
    assert "一括統合できる安全な候補はありませんでした" in r.content.decode()


def test_統合対象が足りなければ断る(as_owner):
    member_gate.切り替える("server")
    r = as_owner.post("/manage/members/duplicates/merge/", {"targetMemberId": "MYM-0001"}, follow=True)
    assert "統合対象が不足しています" in r.content.decode()


# ---- バックアップ / ゴミ箱 ----

def _trash_rows():
    now = timezone.now()
    News.objects.create(sheet_row=2, title="消したNEWS", deleted=True, deleted_at=now - timedelta(days=1), delete_reason="間違い")
    Product.objects.create(sheet_row=3, name="消した商品", deleted=True, deleted_at=now - timedelta(days=2))
    CalendarEvent.objects.create(sheet_row=4, title="消したイベント", deleted=True, deleted_at=now - timedelta(days=3))
    Menu.objects.create(sheet_row=5, name="消したメニュー", deleted=True, deleted_at=now)
    PushNotice.objects.create(sheet_row=6, title="消した通知", deleted=True, deleted_at=now - timedelta(hours=1))
    _member("MYM-0009", "退会 さん", deleted=True, deleted_at=now - timedelta(days=4))
    News.objects.create(sheet_row=7, title="生きているNEWS")


def test_ゴミ箱の一覧と種類と検索(as_owner):
    _trash_rows()
    page = as_owner.get("/manage/trash/").content.decode()
    for label in ["🛟 バックアップ / ゴミ箱", "削除日時", "種類", "内容", "理由", "操作", "一括復元", "一括完全削除", "復元", "完全削除",
                  "会員名・注文ID・タイトルで検索", "全ての種類", "バックアップ状況はシステム管理へ", "/manage/system/"]:
        assert label in page, label
    for s in ["会員", "注文", "NEWS", "ショップ", "カレンダー", "ホーム", "Push通知"]:
        assert f'<option value="{s}"' in page, s
    for t in ["消したNEWS", "消した商品", "消したイベント", "消したメニュー", "消した通知", "退会 さん", "間違い"]:
        assert t in page, t
    assert "生きているNEWS" not in page
    # 新しい順（ホーム → Push通知 → NEWS → ショップ → カレンダー → 会員）
    順 = [x["source"] for x in trash.一覧()["items"]]
    assert 順 == ["ホーム", "Push通知", "NEWS", "ショップ", "カレンダー", "会員"]
    page = as_owner.get("/manage/trash/?source=ショップ").content.decode()
    assert "消した商品" in page and "消したNEWS" not in page
    page = as_owner.get("/manage/trash/?q=イベント").content.decode()
    assert "消したイベント" in page and "消した商品" not in page
    page = as_owner.get("/manage/trash/?q=ない").content.decode()
    assert "条件に一致するゴミ箱項目はありません" in page


def test_ゴミ箱が空のときの文言(as_owner):
    page = as_owner.get("/manage/trash/").content.decode()
    assert "ゴミ箱は空です" in page


def test_復元すると印が外れて一覧に戻る(as_owner):
    _trash_rows()
    r = as_owner.post("/manage/trash/restore/", {"items": ["BLOG:2"]}, follow=True)
    assert "復元しました" in r.content.decode()
    n = News.objects.get(sheet_row=2)
    assert n.deleted is False and n.deleted_at is None and n.delete_reason == ""
    assert "消したNEWS" not in as_owner.get("/manage/trash/").content.decode()
    # まとめて戻す
    r = as_owner.post("/manage/trash/restore/", {"items": ["PRODUCTS:3", "CALENDAR:4", "MENUS:5", "PUSH:6"], "bulk": "1"}, follow=True)
    assert "一括復元しました" in r.content.decode()
    assert Product.objects.get(sheet_row=3).deleted is False and PushNotice.objects.get(sheet_row=6).deleted_at is None
    assert [x["source"] for x in trash.一覧()["items"]] == ["会員"]


def test_完全削除は本当に消える(as_owner):
    _trash_rows()
    r = as_owner.post("/manage/trash/hard-delete/", {"items": ["BLOG:2"]}, follow=True)
    assert "完全削除しました" in r.content.decode()
    assert not News.objects.filter(sheet_row=2).exists()
    r = as_owner.post("/manage/trash/hard-delete/", {"items": ["PRODUCTS:3", "MENUS:5"], "bulk": "1"}, follow=True)
    assert "一括完全削除しました" in r.content.decode()
    assert not Product.objects.filter(sheet_row=3).exists() and not Menu.objects.filter(sheet_row=5).exists()
    assert News.objects.filter(sheet_row=7).exists()  # 生きている行は触らない


def test_選んでいなければ断る(as_owner):
    r = as_owner.post("/manage/trash/restore/", {}, follow=True)
    assert "復元する項目を選択してください" in r.content.decode()
    r = as_owner.post("/manage/trash/hard-delete/", {}, follow=True)
    assert "完全削除する項目を選択してください" in r.content.decode()
    r = as_owner.post("/manage/trash/restore/", {"items": ["BLOG:999"]}, follow=True)
    assert "復元に失敗しました" in r.content.decode()


def test_切り替え前は会員の復元と完全削除を断り会員以外は通す(as_owner):
    _trash_rows()
    r = as_owner.post("/manage/trash/hard-delete/", {"items": ["USERS:MYM-0009"]}, follow=True)
    assert member_gate.断る文() in r.content.decode()
    assert Member.objects.filter(pk="MYM-0009").exists()
    r = as_owner.post("/manage/trash/restore/", {"items": ["USERS:MYM-0009"]}, follow=True)
    assert member_gate.断る文() in r.content.decode()
    assert Member.objects.get(pk="MYM-0009").deleted is True
    # 会員と NEWS を一緒に選ぶと、NEWS だけ戻る
    r = as_owner.post("/manage/trash/restore/", {"items": ["USERS:MYM-0009", "BLOG:2"], "bulk": "1"}, follow=True)
    html = r.content.decode()
    assert member_gate.断る文() in html and "一括復元しました" in html
    assert News.objects.get(sheet_row=2).deleted is False


def test_切り替え後は会員も復元と完全削除ができる(as_owner):
    member_gate.切り替える("server")
    _trash_rows()
    Member.objects.filter(pk="MYM-0009").update(merged_into_id="MYM-0001")
    r = as_owner.post("/manage/trash/restore/", {"items": ["USERS:MYM-0009"]}, follow=True)
    assert "復元しました" in r.content.decode()
    m = Member.objects.get(pk="MYM-0009")
    assert m.deleted is False and m.deleted_at is None and m.merged_into_id == ""
    Member.objects.filter(pk="MYM-0009").update(deleted=True)
    r = as_owner.post("/manage/trash/hard-delete/", {"items": ["USERS:MYM-0009"]}, follow=True)
    assert "完全削除しました" in r.content.decode()
    assert not Member.objects.filter(pk="MYM-0009").exists()


def test_gasapiの窓口からも戻すと完全に消すができる(client, settings):
    """GAS が転送してくる形（restoreDeletedRecord / hardDeleteRecord / getAdminTrashItems）。"""
    import json

    settings.API_KEY = "k"
    _trash_rows()
    r = client.get("/api?action=getAdminTrashItems", HTTP_X_API_KEY="k")
    assert r.status_code == 200 and len(r.json()["items"]) == 6
    r = client.post("/api", json.dumps({"type": "restoreDeletedRecord", "sheet": "BLOG", "rowIdx": 2}),
                    content_type="application/json", HTTP_X_API_KEY="k")
    assert r.json()["status"] == "ok" and News.objects.get(sheet_row=2).deleted is False
    r = client.post("/api", json.dumps({"type": "hardDeleteRecord", "sheet": "PRODUCTS", "rowIdx": 3}),
                    content_type="application/json", HTTP_X_API_KEY="k")
    assert r.json()["status"] == "ok" and not Product.objects.filter(sheet_row=3).exists()
    r = client.post("/api", json.dumps({"type": "hardDeleteRecord", "sheet": "ORDERS", "rowIdx": 3}),
                    content_type="application/json", HTTP_X_API_KEY="k")
    assert r.json()["status"] == "error"
