"""システム管理。**中身は旧管理アプリ（#page-system-management）と同じ**
（集計カード4つ・アラート・未受取特典と未公開予約の明細・公開予約一覧・バックアップ状況）。"""

import os
import subprocess
from datetime import timedelta

import pytest
from django.utils import timezone

from apps.content.models import CalendarEvent, Menu, News, Product, PushNotice
from apps.manage import views_system
from apps.members.models import Member
from apps.records.models import BackupRecord, OrderLine

pytestmark = pytest.mark.django_db


def _member(mid, name, phone="", birthday=None, rewards=None, deleted=False):
    return Member.objects.create(member_id=mid, name=name, phone=phone, birthday=birthday,
                                 reward_history=rewards or [], deleted=deleted)


def test_まゆみ以外は入れない(client, staff):
    client.force_login(staff)
    assert client.get("/manage/system/").status_code == 403


def test_何も無いときの文言(as_owner):
    page = as_owner.get("/manage/system/").content.decode()
    for label in ["会員数", "公開予約", "重複候補", "Push失敗", "🛠️ システム管理", "集計時点"]:
        assert label in page, label
    assert "現在のアラートはありません。" in page
    assert "公開予約はありません。" in page
    # サーバーには控えの置き場が無い → 分からないと正直に出す（GAS のトリガーはサーバーから見えない）
    assert "サーバーからは確認できません。旧管理アプリで確認してください。" in page
    assert "バックアップが古くなっています" not in page
    # pg_dump が無い（テストは SQLite）→ 押せるボタンではなく、その旨の文言
    assert "system/backup/" not in page and "今すぐバックアップ: " in page
    assert "システム管理" in page  # サイドバー


def test_集計カードとアラートは旧アプリと同じ定義で数える(as_owner):
    now = timezone.now()
    # 会員: 退会は数えない。電話番号が同じ2人 → 重複候補1組
    _member("MYM-0001", "山田 花子", phone="090-1111-2222")
    _member("MYM-0002", "山田花子", phone="09011112222")
    _member("MYM-0003", "退会 さん", deleted=True)
    # 30日以上前に獲得して未使用の特典 → 未受取特典
    _member("MYM-0004", "佐藤 桃子", rewards=[
        {"rewardName": "A賞プレゼント", "earnedDate": (now - timedelta(days=40)).strftime("%Y/%m/%d %H:%M"), "used": False},
        {"rewardName": "B賞プレゼント", "earnedDate": (now - timedelta(days=40)).strftime("%Y/%m/%d %H:%M"), "used": True},
        {"rewardName": "C賞プレゼント", "earnedDate": (now - timedelta(days=3)).strftime("%Y/%m/%d %H:%M"), "used": False},
    ])
    # 受付中の注文が2件、うち1件は3日以上前 → 未対応注文2件・長期間未対応1件
    OrderLine.objects.create(order_id="ORD-1", ordered_at=now - timedelta(days=5), status="受付中", product_name="茶", quantity=1)
    OrderLine.objects.create(order_id="ORD-2", ordered_at=now - timedelta(hours=1), status="", product_name="茶", quantity=1)
    OrderLine.objects.create(order_id="ORD-3", ordered_at=now - timedelta(days=9), status="受取済", received=True, product_name="茶", quantity=1)
    # Push失敗（消した分は数えない）
    PushNotice.objects.create(sheet_row=2, title="a", status="送信失敗")
    PushNotice.objects.create(sheet_row=3, title="b", status="送信失敗", deleted=True)
    PushNotice.objects.create(sheet_row=4, title="c", status="自動送信済み")
    # 在庫警告: 売切1つ・閾値以下1つ・ふつう1つ
    Product.objects.create(sheet_row=2, name="売切の品", sold_out="売切", published=True)
    Product.objects.create(sheet_row=3, name="残り少ない品", stock=2, stock_warning=3, published=True)
    Product.objects.create(sheet_row=4, name="ふつうの品", stock=10, stock_warning=3, published=True)
    # 公開予約: 未来のものだけ数える。消したものは数えない
    News.objects.create(sheet_row=2, title="来週のお知らせ", publish_at=now + timedelta(days=7))
    CalendarEvent.objects.create(sheet_row=2, title="", publish_at=now + timedelta(days=2))
    Product.objects.create(sheet_row=5, name="消した品", publish_at=now + timedelta(days=1), deleted=True)
    # 未公開予約: 予定を24時間以上過ぎて非公開のまま
    Menu.objects.create(sheet_row=2, name="遅れているメニュー", publish_at=now - timedelta(days=2), published=False)
    Menu.objects.create(sheet_row=3, name="公開済みメニュー", publish_at=now - timedelta(days=2), published=True)

    page = as_owner.get("/manage/system/").content.decode()
    # 集計カード（順に 会員数3 / 公開予約2 / 重複候補1 / Push失敗1）
    for label, value in [("会員数", 3), ("公開予約", 2), ("重複候補", 1), ("Push失敗", 1)]:
        assert f'<div class="stat-label">{label}</div><div class="stat-value">{value}</div>' in page, label
    # アラートの文言（GAS の buildAdminDashboardData_ と同じ）
    for text in ["2件の受付中注文があります", "1件の送信失敗があります", "2件の商品で在庫警告があります",
                 "1組の重複候補があります", "1件の受付中注文が3日以上経過しています",
                 "1件の特典が30日以上未受取です", "1件の公開予約が予定時刻を過ぎても非公開のままです"]:
        assert text in page, text
    assert "現在のアラートはありません。" not in page
    # 明細
    assert "未受取特典" in page and "佐藤 桃子 (MYM-0004) / A賞プレゼント" in page
    assert "B賞プレゼント" not in page and "C賞プレゼント" not in page
    assert "未公開予約" in page and "ホーム / 遅れているメニュー" in page and "公開済みメニュー" not in page
    # 公開予約一覧は近い順。題名が空のカレンダーは「イベント」
    assert page.index("カレンダー / イベント") < page.index("NEWS / 来週のお知らせ")
    assert "消した品" not in page


def test_重複の組は電話と氏名と生年月日で作り同じ顔ぶれは1組(as_owner):
    from datetime import date

    会員 = [
        {"memberId": "A", "name": "山田 花子", "phone": "9011112222", "birthday": "2000-01-02"},
        {"memberId": "B", "name": "山田花子", "phone": "090-1111-2222", "birthday": "2000/1/2"},
        {"memberId": "C", "name": "鈴木 一郎", "phone": "", "birthday": ""},
        {"memberId": "D", "name": "鈴木一郎", "phone": "080-0000-0000", "birthday": ""},
        {"memberId": "E", "name": "", "phone": "", "birthday": str(date(1990, 1, 1))},
    ]
    組 = views_system._重複の組(会員)
    assert [sorted(u["memberId"] for u in g) for g in 組] == [["A", "B"], ["C", "D"]]


def test_控えの置き場があれば最終日時と古さを出す(as_owner, tmp_path, monkeypatch):
    monkeypatch.setenv("BACKUP_DIR", str(tmp_path))
    page = as_owner.get("/manage/system/").content.decode()
    assert "最終バックアップ: 未実行" in page and "バックアップが古くなっています" in page
    assert "36時間以上バックアップが更新されていません" in page

    古い = tmp_path / "mayumi-20260901-0300.dump"
    古い.write_bytes(b"x" * 2048)
    古い_t = (timezone.now() - timedelta(days=3)).timestamp()
    os.utime(古い, (古い_t, 古い_t))
    page = as_owner.get("/manage/system/").content.decode()
    assert "バックアップが古くなっています" in page and "mayumi-20260901-0300.dump" in page

    新しい = tmp_path / "mayumi-20260915-0300.dump"
    新しい.write_bytes(b"x" * 2048)
    page = as_owner.get("/manage/system/").content.decode()
    assert "バックアップが古くなっています" not in page
    assert "36時間以上バックアップが更新されていません" not in page
    assert "最終バックアップ: " + timezone.localtime().strftime("%Y/%m/%d") in page
    assert "最近のバックアップ履歴" in page and "mayumi-20260915-0300.dump" in page


def test_スプレッドシートの控えの記録は取り込み時点までと断って出す(as_owner):
    BackupRecord.objects.create(sheet_row=2, created_at=timezone.now(), kind="manual-admin",
                                file_name="まゆみ助産院_管理_backup_20260405", url="https://drive.google.com/x")
    page = as_owner.get("/manage/system/").content.decode()
    assert "スプレッドシートの控えの記録" in page and "取り込んだ時点まで" in page
    assert 'href="https://drive.google.com/x"' in page and "manual-admin" in page


def test_今すぐバックアップはpg_dumpが無ければ取れない理由を返す(as_owner):
    r = as_owner.post("/manage/system/backup/", follow=True)
    assert r.redirect_chain[-1][0].endswith("/manage/system/")
    assert "サーバーからは控えを取れません" in r.content.decode()


def test_今すぐバックアップはpg_dumpがあれば控えを1つ書く(as_owner, tmp_path, monkeypatch):
    monkeypatch.setenv("BACKUP_DIR", str(tmp_path / "backups"))
    monkeypatch.setattr(views_system, "_接続の設定", lambda: {
        "ENGINE": "django.db.backends.postgresql", "NAME": "mayumi", "USER": "postgres",
        "PASSWORD": "pw", "HOST": "db", "PORT": "5432"})
    monkeypatch.setattr(views_system, "_pg_dumpの場所", lambda: "/usr/bin/pg_dump")
    呼んだ = {}

    def fake_run(cmd, env=None, **kw):
        呼んだ["cmd"] = cmd
        呼んだ["pw"] = env.get("PGPASSWORD")
        # -f の次がファイル。空の控えを「取れた」と言わないよう 1KB より大きく書く
        with open(cmd[cmd.index("-f") + 1], "wb") as f:
            f.write(b"x" * 4096)
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(views_system.subprocess, "run", fake_run)
    page = as_owner.get("/manage/system/").content.decode()
    assert "system/backup/" in page and "💾 今すぐバックアップ" in page
    r = as_owner.post("/manage/system/backup/", follow=True)
    assert "バックアップを作成しました" in r.content.decode()
    assert 呼んだ["pw"] == "pw" and "-Fc" in 呼んだ["cmd"] and 呼んだ["cmd"][-1] == "mayumi"
    files = list((tmp_path / "backups").glob("mayumi-*.dump"))
    assert len(files) == 1

    # 中身が空なら失敗として消す（気づかず何日も過ぎるのが一番困る）
    def empty_run(cmd, env=None, **kw):
        open(cmd[cmd.index("-f") + 1], "wb").close()
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(views_system.subprocess, "run", empty_run)
    monkeypatch.setenv("BACKUP_DIR", str(tmp_path / "backups2"))
    r = as_owner.post("/manage/system/backup/", follow=True)
    assert "バックアップに失敗しました" in r.content.decode()
    assert list((tmp_path / "backups2").glob("mayumi-*.dump")) == []
