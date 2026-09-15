"""お客様アプリへの通知（OneSignal）を、この管理画面から送る。

GAS の sendAutoPush と同じ内容（タグで全員・開くページ）。
**設定が無ければ送らない**、送れなくても投稿は残る。
"""

import io
import json

import pytest
from django.utils import timezone

from apps.content.models import Category, News, PushNotice
from apps.content import onesignal
from apps.manage import push  # noqa: F401

pytestmark = pytest.mark.django_db


@pytest.fixture
def keys(settings):
    settings.ONESIGNAL_APP_ID = "app-0001"
    settings.ONESIGNAL_REST_API_KEY = "key-0001"


@pytest.fixture
def fake_onesignal(monkeypatch):
    """OneSignal へ実際には行かない。送った中身を控える。"""
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append({"url": req.full_url, "headers": dict(req.header_items()), "body": json.loads(req.data)})
        return _Res(json.dumps({"id": "n-123", "recipients": 42}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


def _form(**extra):
    Category.objects.get_or_create(sheet_row=2, name="イベント情報", defaults={"kind": "お知らせ"})
    data = {
        "posted_on": "2026-09-14", "title": "マッサージ教室", "category": "イベント情報", "icon": "📢",
        "body": "本文", "publish_at": "", "link_url": "", "button_text": "", "notice_listed": "on",
        "status": "公開",
    }
    data.update(extra)
    return data


def test_設定が無ければ送らず投稿は残る(as_owner, settings):
    settings.ONESIGNAL_APP_ID = ""
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert News.objects.filter(title="マッサージ教室").exists()
    assert PushNotice.objects.count() == 0


def test_チェックすると全員へ送り記録が残る(as_owner, keys, fake_onesignal):
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert len(fake_onesignal) == 1
    body = fake_onesignal[0]["body"]
    assert body["app_id"] == "app-0001"
    assert body["headings"]["ja"] == "📝 マッサージ教室"
    assert body["contents"]["ja"] == "NEWSが更新されました"
    assert body["filters"] == [{"field": "tag", "key": "app_scope", "relation": "=", "value": "mayumi_josanin_app"}]
    assert body["url"].endswith("?open=news") and body["data"]["openPage"] == "news"
    assert fake_onesignal[0]["headers"]["Authorization"] == "Basic key-0001"

    rec = PushNotice.objects.get()
    assert rec.status == "自動送信済み"
    assert rec.notification_id == "n-123"
    assert rec.recipient_count == 42
    assert rec.sent_at is not None
    assert rec.sheet_row == 2


def test_チェックしなければ送らない(as_owner, keys, fake_onesignal):
    as_owner.post("/manage/news/new/", _form())
    assert fake_onesignal == []


def test_下書きや公開開始が先なら送らない(as_owner, keys, fake_onesignal):
    as_owner.post("/manage/news/new/", _form(send_push="on", status="非公開"))
    future = (timezone.localtime() + timezone.timedelta(days=2)).strftime("%Y-%m-%dT%H:%M")
    as_owner.post("/manage/news/new/", _form(send_push="on", title="先の予定", publishAt=future))
    assert fake_onesignal == []


def test_送れなくても投稿は残り失敗が記録される(as_owner, keys, monkeypatch):
    def boom(req, timeout=0):
        raise onesignal.urllib.error.URLError("down")

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", boom)
    r = as_owner.post("/manage/news/new/", _form(send_push="on"))
    assert r.status_code == 302
    assert News.objects.filter(title="マッサージ教室").exists()
    rec = PushNotice.objects.get()
    assert rec.status == "送信失敗"


# ========== 管理画面（旧管理アプリの「📣 全員へPush通知を送信」「📣 送信済みお知らせ（Push通知）履歴」と同じ中身） ==========

from datetime import timedelta  # noqa: E402

from apps.gasapi import admin_push  # noqa: E402
from apps.members.models import Member  # noqa: E402
from apps.records.models import OrderLine  # noqa: E402


def _member(member_id, name, sub):
    return Member.objects.create(member_id=member_id, name=name, push_subscription=sub)


def test_送信フォームの項目は旧管理アプリと同じ(as_owner, keys):
    _member("MYM-0001", "山田", "sub-abc")
    page = as_owner.get("/manage/push/").content.decode()
    for label in ["📣 全員へPush通知を送信", "Push通知を許可しているすべてのユーザーのスマートフォンに通知を送信します。",
                  "通知タイトル", "通知本文", "配信対象", "開く画面", "配信予定日時", "テスト送信先",
                  "👁 プレビュー", "📁 下書き保存", "🧪 テスト送信", "🕒 予約保存", "🚀 通知を一斉送信する",
                  'value="まゆみ助産院"', "例：新しいブログを公開しました！", "通知の内容を入力してください..."]:
        assert label in page, label
    for opt in ["Push許可中の全員", "受付中注文がある会員", "未使用特典がある会員", "注文履歴がない会員"]:
        assert opt in page, opt
    for opt in ['value="home">ホーム', 'value="shop">ショップ', 'value="calendar">カレンダー', 'value="news">NEWS',
                'value="notices">お知らせ一覧', 'value="mypage">マイページ', 'value="menu-list">メニュー一覧']:
        assert opt in page, opt
    assert "会員を選択してください" in page and "山田 (MYM-0001)" in page
    assert "送り直す" not in page  # 旧管理アプリに無い


def test_履歴の集計カードと検索と空の文言(as_owner, keys):
    page = as_owner.get("/manage/push/").content.decode()
    for label in ["📣 送信済みお知らせ（Push通知）履歴", "🗑 一括削除", "🔄 更新", "通知総数", "7日以内", "現在表示中",
                  "タイトル・本文で検索", "条件をクリア", "通知履歴がありません", "通知履歴はありません",
                  "<th>送信日時</th><th>タイトル</th><th>本文</th><th>操作</th>"]:
        assert label in page, label

    admin_push.送る({"mode": "draft", "title": "秋の教室", "body": "募集中"})
    admin_push.送る({"mode": "draft", "title": "年末のご案内", "body": "お休みです"})
    page = as_owner.get("/manage/push/").content.decode()
    assert "2 / 2 件を表示" in page and "下書き / home / all" in page
    # 新しい順（あとで書いた方が上）
    assert page.index("年末のご案内") < page.index("秋の教室")
    page = as_owner.get("/manage/push/?q=お休み").content.decode()
    assert "年末のご案内" in page and "秋の教室" not in page and "1 / 2 件を表示" in page
    page = as_owner.get("/manage/push/?q=無い言葉").content.decode()
    assert "条件に一致する通知履歴はありません" in page


def test_7日以内の数は送った日時で数える(as_owner, keys):
    admin_push.送る({"mode": "draft", "title": "古い", "body": "x"})
    old = PushNotice.objects.get()
    old.updated_at = timezone.now() - timedelta(days=10)
    old.save(update_fields=["updated_at"])
    admin_push.送る({"mode": "draft", "title": "新しい", "body": "x"})
    page = as_owner.get("/manage/push/").content.decode()
    i = page.index("7日以内")
    assert '<div class="stat-value">1</div>' in page[i:i + 200]


def test_一斉送信は全員宛てで履歴に残る(as_owner, keys, fake_onesignal):
    r = as_owner.post("/manage/push/send/", {"mode": "send", "title": "教室のご案内", "body": "本文", "targetPage": "news",
                                           "targetStatus": "all"}, follow=True)
    page = r.content.decode()
    assert "送信完了しました！" in page
    body = fake_onesignal[0]["body"]
    assert "filters" in body and "include_subscription_ids" not in body and body["data"]["openPage"] == "news"
    assert "送信済み / news / all" in page and "結果: 送信済み" in page and "対象件数: 42件" in page


def test_配信対象で絞ると届け先だけに送る(as_owner, keys, fake_onesignal):
    _member("MYM-0001", "山田", "sub-abc")
    _member("MYM-0002", "佐藤", "sub-def")
    OrderLine.objects.create(sheet_row=2, order_id="ORD-1", member_id="MYM-0002", status="受付中")
    r = as_owner.post("/manage/push/send/", {"mode": "send", "title": "t", "body": "b", "targetStatus": "no-orders"}, follow=True)
    assert "送信完了しました！" in r.content.decode()
    assert fake_onesignal[0]["body"]["include_subscription_ids"] == ["sub-abc"]
    p = PushNotice.objects.get()
    assert p.target_status == "no-orders" and p.target_detail == "注文履歴がない会員"


def test_テスト送信は選んだ会員だけ(as_owner, keys, fake_onesignal):
    _member("MYM-0001", "山田", "sub-abc")
    r = as_owner.post("/manage/push/send/", {"mode": "test", "title": "試し", "body": "本文"}, follow=True)
    assert "テスト送信先を選択してください" in r.content.decode() and fake_onesignal == []
    r = as_owner.post("/manage/push/send/", {"mode": "test", "title": "試し", "body": "本文", "testRecipient": "MYM-0001"}, follow=True)
    assert "テスト通知を送信しました" in r.content.decode()
    assert fake_onesignal[0]["body"]["include_subscription_ids"] == ["sub-abc"]
    assert PushNotice.objects.get().result == "テスト送信済み"


def test_下書きと予約(as_owner, keys, fake_onesignal):
    r = as_owner.post("/manage/push/send/", {"mode": "draft", "title": "まゆみ助産院", "body": ""}, follow=True)
    assert "タイトルと本文を入力してください" in r.content.decode() and PushNotice.objects.count() == 0
    r = as_owner.post("/manage/push/send/", {"mode": "draft", "title": "下書き", "body": "b"}, follow=True)
    assert "下書きを保存しました" in r.content.decode() and PushNotice.objects.get().status == "下書き"
    r = as_owner.post("/manage/push/send/", {"mode": "schedule", "title": "明日", "body": "b"}, follow=True)
    assert "タイトル・本文・配信予定日時を入力してください" in r.content.decode()
    when = timezone.localtime() + timedelta(days=1)
    r = as_owner.post("/manage/push/send/", {"mode": "schedule", "title": "明日", "body": "b",
                                           "scheduledAt": when.strftime("%Y-%m-%dT%H:%M")}, follow=True)
    page = r.content.decode()
    assert "通知を予約しました" in page and "予約: " + when.strftime("%Y/%m/%d") in page and "予約済み / home / all" in page
    assert fake_onesignal == []


def test_削除と一括削除(as_owner, keys):
    for t in ("a", "b", "c"):
        admin_push.送る({"mode": "draft", "title": t, "body": "x"})
    rows = list(PushNotice.objects.order_by("sheet_row").values_list("sheet_row", flat=True))
    r = as_owner.post("/manage/push/bulk-delete/", {}, follow=True)
    assert "削除する項目を選択してください。" in r.content.decode()
    r = as_owner.post("/manage/push/bulk-delete/", {"rows": [str(rows[0]), str(rows[1])]}, follow=True)
    assert "2件の通知履歴を削除しました" in r.content.decode()
    r = as_owner.post(f"/manage/push/{rows[2]}/delete/", follow=True)
    assert "通知を削除しました" in r.content.decode()
    assert PushNotice.objects.filter(deleted=False).count() == 0
    assert "通知履歴がありません" in as_owner.get("/manage/push/").content.decode()
