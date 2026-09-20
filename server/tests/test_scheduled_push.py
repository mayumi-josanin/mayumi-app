"""予約したお知らせ通知を、サーバー（cron の箱）が時間どおりに送る。

2026-09-20 まで、時間が来たときに送る役目は GAS の自動実行だけが担っていた。
その GAS が2日間で6回落ちていたので、サーバー側にも同じ役目を持たせた。
**いちばん大事なのは「二度送らない」こと。**
"""

import io
import json
from datetime import timedelta

import pytest
from django.core.management import call_command
from django.utils import timezone

from apps.content import onesignal
from apps.content.models import PushNotice
from apps.members.models import Member
from apps.records.models import OrderLine

pytestmark = pytest.mark.django_db


@pytest.fixture
def keys(settings):
    settings.ONESIGNAL_APP_ID = "app-0001"
    settings.ONESIGNAL_REST_API_KEY = "key-0001"


@pytest.fixture
def 送った中身(monkeypatch):
    """OneSignal へは実際に行かない。送った中身を控える。"""
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append(json.loads(req.data))
        return _Res(json.dumps({"id": "n-555", "recipients": 12}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


def _予約(title, *, 分前=10, status="予約済み", 日時あり=True, target_status="all", row=None):
    予定 = (timezone.now() - timedelta(minutes=分前)) if 日時あり else None
    行 = row if row is not None else (PushNotice.objects.count() + 2)
    return PushNotice.objects.create(
        sheet_row=行, title=title, body="本文", status=status,
        scheduled_at=予定, target_page="news", target_status=target_status,
    )


def _流す(*args):
    出 = io.StringIO()
    call_command("予約した通知を送る", *args, stdout=出, stderr=出)
    return 出.getvalue()


def test_時間が来た予約は送られ履歴に残る(keys, 送った中身):
    p = _予約("そろそろ")
    _流す()
    assert len(送った中身) == 1
    assert 送った中身[0]["headings"]["ja"] == "そろそろ"
    assert 送った中身[0]["data"]["openPage"] == "news"
    p.refresh_from_db()
    assert p.status == "送信済み" and p.sent_at is not None
    assert p.result == "予約送信済み" and p.notification_id == "n-555"
    assert p.recipient_count == 12


def test_まだ時間が来ていないものは送らない(keys, 送った中身):
    p = _予約("まだ", 分前=-60)  # 1時間後
    _流す()
    assert 送った中身 == []
    p.refresh_from_db()
    assert p.status == "予約済み" and p.sent_at is None


def test_下書きと予約日時が無いものと送信済みは触らない(keys, 送った中身):
    下書き = _予約("下書き", status="下書き")
    日時なし = _予約("日時なし", 日時あり=False)
    済み = _予約("済み", status="送信済み")
    済み.sent_at = timezone.now()
    済み.save()
    _流す()
    assert 送った中身 == []
    for p in (下書き, 日時なし, 済み):
        p.refresh_from_db()
    assert 下書き.status == "下書き" and 日時なし.status == "予約済み"
    assert 日時なし.sent_at is None


def test_二度流しても二度送らない(keys, 送った中身):
    _予約("一度だけ")
    _流す()
    _流す()
    assert len(送った中身) == 1


def test_送信に失敗したら印を戻して次へ進む(keys, monkeypatch):
    駄目 = _予約("届かない", row=2)
    平気 = _予約("届く", row=3)

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        # 送る中身は ASCII に逃がして送られるので、読み解いてから題を見る
        if json.loads(req.data)["headings"]["ja"] == "届かない":
            raise onesignal.urllib.error.URLError("down")
        return _Res(json.dumps({"id": "n-9", "recipients": 3}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    _流す()
    駄目.refresh_from_db()
    平気.refresh_from_db()
    # 印（送った日時）は戻す。1件の失敗で残りを止めない。
    assert 駄目.status == "送信失敗" and 駄目.sent_at is None and "down" in 駄目.result
    assert 平気.status == "送信済み" and 平気.result == "予約送信済み"


def test_失敗したものは次の1分でやり直さない(keys, monkeypatch):
    """毎分やり直すと、届かないまま何度も試し続けることになる。履歴に残して院長が送り直す。"""
    _予約("届かない")

    def boom(req, timeout=0):
        raise onesignal.urllib.error.URLError("down")

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", boom)
    _流す()
    回数 = []
    monkeypatch.setattr(onesignal.urllib.request, "urlopen",
                        lambda req, timeout=0: 回数.append(1))
    _流す()
    assert 回数 == []


def test_下見は送らずに件数と中身だけ出す(keys, 送った中身):
    p = _予約("下見するもの")
    出 = _流す("--下見")
    assert 送った中身 == []
    assert "送る予約 1件" in 出 and "下見するもの" in 出
    p.refresh_from_db()
    assert p.status == "予約済み" and p.sent_at is None


def test_下見で対象が無ければそう言う(keys):
    assert "送る予約はありません" in _流す("--下見")


def test_鍵が無いときは落ちずに記録だけ残す(settings, 送った中身):
    settings.ONESIGNAL_APP_ID = ""
    settings.ONESIGNAL_REST_API_KEY = ""
    p = _予約("鍵が無い")
    出 = _流す()
    assert 送った中身 == []
    assert "通知の鍵が設定されていないため送れませんでした" in 出
    p.refresh_from_db()
    # 状態は「予約済み」のまま。鍵を入れれば次の1分で送られる。
    assert p.status == "予約済み" and p.sent_at is None
    assert "通知の鍵が設定されていないため送れませんでした" in p.result


def test_送り先の絞り込みは画面と同じ(keys, 送った中身):
    Member.objects.create(member_id="MYM-0001", name="山田", push_subscription="sub-abc")
    Member.objects.create(member_id="MYM-0002", name="佐藤", push_subscription="sub-def")
    OrderLine.objects.create(sheet_row=2, order_id="ORD-1", member_id="MYM-0002", status="受付中")
    _予約("注文のない方へ", target_status="no-orders")
    _流す()
    # 画面から「注文履歴がない会員」で送ったときと同じ届け先（tests/test_manage_push.py と同じ）
    assert 送った中身[0]["include_subscription_ids"] == ["sub-abc"]


def test_全員のときは端末を指定しない(keys, 送った中身):
    Member.objects.create(member_id="MYM-0001", name="山田", push_subscription="sub-abc")
    _予約("全員へ")
    _流す()
    assert "include_subscription_ids" not in 送った中身[0]
    assert 送った中身[0]["filters"][0]["value"] == "mayumi_josanin_app"
