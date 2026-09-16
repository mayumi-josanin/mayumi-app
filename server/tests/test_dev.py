"""開発（apps/dev）。**KEM_DDENKI の開発管理の写し**（プロジェクト／タスク／コメント／目安箱、Discord 通知）。

画面の見出し・列・ボタン・文言が KEM と同じであること、まゆみだけが入れること、
Discord は URL があるときだけ送り、送れなくても保存が成立することを確かめる。
"""

import datetime
import io
import json
import re

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.dev import notifications
from apps.dev.models import DevComment, DevProject, DevTask, Meyasubako

pytestmark = pytest.mark.django_db


@pytest.fixture
def project(owner):
    return DevProject.objects.create(name="予約システム", status="in_progress", assignee=owner,
                                     start_date=datetime.date(2026, 9, 1), due_date=datetime.date(2026, 9, 30))


@pytest.fixture
def task(project, owner):
    return DevTask.objects.create(project=project, title="二重予約を止める", category="bug", priority="high", assignee=owner)


@pytest.fixture
def discord(monkeypatch):
    """Discord への POST を捕まえる。送った payload の一覧を返す。"""
    sent = []

    def fake_urlopen(req, timeout=0):
        sent.append({"url": req.full_url, "body": json.loads(req.data)})
        return io.BytesIO(b"")

    monkeypatch.setattr(notifications, "urlopen", fake_urlopen)
    return sent


def _png():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (2400, 1200), (80, 120, 200)).save(buf, format="PNG")
    return SimpleUploadedFile("shot.png", buf.getvalue(), content_type="image/png")


# ========== 入れる人 ==========


def test_スタッフは入れない(client, staff):
    client.force_login(staff)
    for url in ["/manage/dev/", "/manage/dev/new/", "/manage/dev/meyasubako/", "/manage/dev/meyasubako/new/"]:
        assert client.get(url).status_code == 403, url


def test_入っていなければログインへ(client):
    r = client.get("/manage/dev/")
    assert r.status_code == 302 and r["Location"].startswith("/manage/login/")


# ========== プロジェクト ==========


def test_一覧の見出しと列と空の文言(as_owner):
    page = as_owner.get("/manage/dev/").content.decode()
    for s in ["プロジェクト一覧", "+ プロジェクト作成", "稼働中", "自分の担当", "すべて",
              "計画中・進行中のプロジェクトを表示しています", "稼働中のプロジェクトがありません", "プロジェクトがありません"]:
        assert s in page, s
    for th in ["<th>プロジェクト名</th>", "<th>ステータス</th>", "<th>責任者</th>", "<th>期間</th>", "<th>進捗</th>", "<th>未解決バグ</th>"]:
        assert th in page, th
    # 開発の topbar と左メニュー
    assert '<span class="topbar-title">開発</span>' in page


def test_作成すると詳細へ飛び一覧とガントに出る(as_owner, owner):
    r = as_owner.post("/manage/dev/new/", {
        "name": "会員移行", "status": "planning", "assignee": owner.pk,
        "start_date": "2026-09-19", "due_date": "2026-09-21", "description": "GAS からサーバーへ",
        "discord_webhook_url": "",
    })
    made = DevProject.objects.get(name="会員移行")
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/{made.pk}/"
    assert made.status == "planning" and made.assignee == owner and made.description == "GAS からサーバーへ"
    assert "プロジェクトを作成しました。" in as_owner.get(r["Location"]).content.decode()
    page = as_owner.get("/manage/dev/").content.decode()
    assert "会員移行" in page and '<span class="badge badge-gray">計画中</span>' in page
    assert "2026年9月19日 ~ 2026年9月21日" in page  # 日付は LANGUAGE_CODE=ja の見た目
    # ガントの元データ（json_script）と frappe-gantt の CDN
    assert 'id="gantt-data"' in page and "frappe-gantt@0.6.1" in page
    assert "bar-dev-planning" in page and "1件を表示しています" in page


def test_日付が無いプロジェクトは推定で出る(as_owner):
    p = DevProject.objects.create(name="未定")
    DevTask.objects.create(project=p, title="a", due_date=datetime.date(2026, 10, 5))
    page = as_owner.get("/manage/dev/").content.decode()
    assert "（推定）" in page and "bar-dev-inferred" in page
    assert "斜線のバー1件は開始日・期限が未設定のため" in page


def test_絞り込み_稼働中と自分の担当とすべて(as_owner, owner, staff):
    DevProject.objects.create(name="動いている", status="in_progress", assignee=staff)
    DevProject.objects.create(name="終わった", status="completed", assignee=owner)
    active = as_owner.get("/manage/dev/").content.decode()
    assert "動いている" in active and "終わった" not in active
    mine = as_owner.get("/manage/dev/?scope=mine").content.decode()
    assert "終わった" in mine and "動いている" not in mine
    assert "自分が責任者のプロジェクトを表示しています" in mine
    everything = as_owner.get("/manage/dev/?scope=all").content.decode()
    assert "動いている" in everything and "終わった" in everything
    # 知らない scope は稼働中に戻す
    assert "動いている" in as_owner.get("/manage/dev/?scope=zzz").content.decode()


def test_詳細_集計とカンバンの列とタスク一覧(as_owner, project, owner):
    DevTask.objects.create(project=project, title="設計", status="done", estimate_hours=2, actual_hours=3)
    DevTask.objects.create(project=project, title="実装", status="in_progress", estimate_hours=5)
    DevTask.objects.create(project=project, title="バグ直し", status="open", category="bug", priority="critical")
    DevTask.objects.create(project=project, title="見てもらう", status="review")
    page = as_owner.get(f"/manage/dev/{project.pk}/").content.decode()
    for s in ["+ タスク追加", "編集", "削除", "一覧に戻る", "プロジェクト情報", "カンバンボード", "タスク一覧",
              "1/4 (25%)", "未解決バグ", "見積合計", "実績合計"]:
        assert s in page, s
    # 見積・実績の合計（SQLite と PostgreSQL で小数の出方が違うので "7h" も "7.0h" も通す）
    assert re.search(r"見積合計</dt><dd>7(\.0)?h", page) and re.search(r"実績合計</dt><dd>3(\.0)?h", page)
    # カンバンの4列と件数
    for col in ["未着手\n          (1)", "作業中\n          (1)", "レビュー\n          (1)", "完了\n          (1)"]:
        assert col in page, col
    assert "なし" not in page.split("カンバンボード")[1].split("タスク一覧")[0]
    # 進捗バー
    assert "width: 25%" in page
    # タスク一覧の列と絞り込みボタン
    for th in ["<th>ID</th>", "<th>タイトル</th>", "<th>優先度</th>", "<th>カテゴリ</th>", "<th>担当者</th>", "<th>期限</th>"]:
        assert th in page, th
    assert '<span class="badge badge-red">緊急</span>' in page and '<span class="badge badge-red">バグ修正</span>' in page
    only_done = as_owner.get(f"/manage/dev/{project.pk}/?status=done").content.decode()
    rows = only_done.split("<tbody>")[1]
    assert "設計" in rows and "実装" not in rows


def test_詳細_タスクが無いときの文言(as_owner, project):
    page = as_owner.get(f"/manage/dev/{project.pk}/").content.decode()
    assert "0/0 (0%)" in page and "タスクがありません" in page and page.count(">なし</div>") == 4


def test_編集(as_owner, project):
    page = as_owner.get(f"/manage/dev/{project.pk}/edit/").content.decode()
    assert "予約システム を編集" in page and 'value="予約システム"' in page and "Discord Webhook URL" in page
    r = as_owner.post(f"/manage/dev/{project.pk}/edit/", {
        "name": "予約システム v2", "status": "completed", "assignee": "", "start_date": "", "due_date": "",
        "description": "", "discord_webhook_url": "https://discord.com/api/webhooks/1/abc",
    })
    assert r.status_code == 302
    project.refresh_from_db()
    assert project.name == "予約システム v2" and project.status == "completed" and project.assignee is None
    assert project.discord_webhook_url == "https://discord.com/api/webhooks/1/abc"


def test_削除_確認画面を経てタスクとコメントも消える(as_owner, project, task, owner):
    DevComment.objects.create(task=task, author=owner, body="メモ")
    page = as_owner.get(f"/manage/dev/{project.pk}/delete/").content.decode()
    assert "「<strong>予約システム</strong>」を削除しますか？関連するタスク・コメントもすべて削除されます。" in page
    assert DevProject.objects.count() == 1  # GET では消えない
    r = as_owner.post(f"/manage/dev/{project.pk}/delete/")
    assert r.status_code == 302 and r["Location"] == "/manage/dev/"
    assert DevProject.objects.count() == 0 and DevTask.objects.count() == 0 and DevComment.objects.count() == 0


# ========== タスク ==========


def test_タスク作成_項目と表示順(as_owner, project, owner):
    page = as_owner.get(f"/manage/dev/{project.pk}/tasks/new/").content.decode()
    for s in ["タスクを作成", "タイトル", "ステータス", "優先度", "カテゴリ", "担当者", "期限", "見積(h)", "詳細"]:
        assert s in page, s
    assert "実績(h)" not in page  # 実績は編集のときだけ
    # 担当者の候補は管理画面のユーザーだけ
    assert f'<option value="{owner.pk}">まゆみ</option>' in page
    r = as_owner.post(f"/manage/dev/{project.pk}/tasks/new/", {
        "title": "枠を15分に", "status": "open", "priority": "medium", "category": "feature", "assignee": "",
        "due_date": "2026-09-20", "estimate_hours": "1.5", "description": "", "github_issue_url": "", "github_pr_url": "",
    })
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/{project.pk}/"
    made = DevTask.objects.get(title="枠を15分に")
    assert made.project == project and made.sort_order == 0 and str(made.estimate_hours) == "1.5"
    as_owner.post(f"/manage/dev/{project.pk}/tasks/new/", {
        "title": "2つ目", "status": "open", "priority": "low", "category": "docs", "assignee": "",
    })
    assert DevTask.objects.get(title="2つ目").sort_order == 1


def test_タスク編集_実績が出る(as_owner, task):
    page = as_owner.get(f"/manage/dev/tasks/{task.pk}/edit/").content.decode()
    assert f"#{task.pk} 二重予約を止める を編集" in page and "実績(h)" in page
    r = as_owner.post(f"/manage/dev/tasks/{task.pk}/edit/", {
        "title": "二重予約を止める", "status": "done", "priority": "high", "category": "bug", "assignee": "",
        "actual_hours": "4", "description": "除外制約で止めた",
    })
    assert r.status_code == 302
    task.refresh_from_db()
    assert task.status == "done" and str(task.actual_hours) == "4.0" and task.assignee is None


def test_タスク移動_カンバンからステータスを変える(as_owner, project, task):
    r = as_owner.post(f"/manage/dev/tasks/{task.pk}/move/", {"status": "review"})
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/{project.pk}/"
    task.refresh_from_db()
    assert task.status == "review"
    # 知らない値は無視
    as_owner.post(f"/manage/dev/tasks/{task.pk}/move/", {"status": "zzz"})
    task.refresh_from_db()
    assert task.status == "review"


def test_タスク詳細とコメント(as_owner, project, task, owner):
    page = as_owner.get(f"/manage/dev/tasks/{task.pk}/").content.decode()
    for s in [f"#{task.pk} 二重予約を止める", "プロジェクトに戻る", "説明なし", "コメント (0)", "コメントなし",
              'placeholder="コメントを入力..."', "情報", '<span class="badge badge-orange">高</span>', "バグ修正", "まゆみ"]:
        assert s in page, s
    r = as_owner.post(f"/manage/dev/tasks/{task.pk}/", {"body": "除外制約で\n止める"})
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/tasks/{task.pk}/"
    c = DevComment.objects.get()
    assert c.task == task and c.author == owner and c.body == "除外制約で\n止める"
    page = as_owner.get(f"/manage/dev/tasks/{task.pk}/").content.decode()
    assert "コメント (1)" in page and "除外制約で<br>止める" in page and "<strong style=\"color: var(--text-primary);\">まゆみ</strong>" in page
    # 空のコメントは入らない
    as_owner.post(f"/manage/dev/tasks/{task.pk}/", {"body": ""})
    assert DevComment.objects.count() == 1


def test_タスク削除(as_owner, project, task, owner):
    DevComment.objects.create(task=task, author=owner, body="メモ")
    page = as_owner.get(f"/manage/dev/tasks/{task.pk}/delete/").content.decode()
    assert f"「<strong>#{task.pk} 二重予約を止める</strong>」を削除しますか？関連するコメントもすべて削除されます。" in page
    r = as_owner.post(f"/manage/dev/tasks/{task.pk}/delete/")
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/{project.pk}/"
    assert DevTask.objects.count() == 0 and DevComment.objects.count() == 0
    assert DevProject.objects.count() == 1


# ========== Discord 通知 ==========


def test_Discord_URLが無ければ送らない(as_owner, project, owner, discord):
    as_owner.post(f"/manage/dev/{project.pk}/tasks/new/", {
        "title": "t", "status": "open", "priority": "medium", "category": "feature", "assignee": owner.pk,
    })
    assert discord == []


def test_Discord_割当と状態変更で送る(as_owner, project, owner, staff, discord):
    project.discord_webhook_url = "https://discord.com/api/webhooks/1/abc"
    project.save()
    as_owner.post(f"/manage/dev/{project.pk}/tasks/new/", {
        "title": "通知を試す", "status": "open", "priority": "critical", "category": "infra", "assignee": staff.pk,
        "due_date": "2026-09-25", "description": "本文",
    })
    assert len(discord) == 1
    sent = discord[0]
    assert sent["url"] == "https://discord.com/api/webhooks/1/abc"
    assert sent["body"]["content"] == "**タスクが割り当てられました**"
    embed = sent["body"]["embeds"][0]
    task = DevTask.objects.get(title="通知を試す")
    assert embed["title"] == f"#{task.pk} 通知を試す" and embed["description"] == "本文"
    fields = {f["name"]: f["value"] for f in embed["fields"]}
    assert fields == {"担当者": "スタッフ", "優先度": "緊急 !!", "カテゴリ": "インフラ", "ステータス": "未着手",
                      "期限": "2026-09-25", "プロジェクト": "予約システム"}
    assert embed["footer"] == {"text": "割当者: まゆみ"}

    # カンバンの移動 → ステータス変更の通知
    as_owner.post(f"/manage/dev/tasks/{task.pk}/move/", {"status": "in_progress"})
    assert len(discord) == 2 and discord[1]["body"]["content"] == "**ステータスが変更されました**"
    fields = {f["name"]: f["value"] for f in discord[1]["body"]["embeds"][0]["fields"]}
    assert fields == {"ステータス変更": "未着手 → 作業中", "担当者": "スタッフ", "変更者": "まゆみ"}

    # 編集で担当者が変わらなければ割当は送らない。ステータスだけ変わればその通知だけ
    as_owner.post(f"/manage/dev/tasks/{task.pk}/edit/", {
        "title": "通知を試す", "status": "done", "priority": "critical", "category": "infra", "assignee": staff.pk,
    })
    assert len(discord) == 3 and discord[2]["body"]["content"] == "**ステータスが変更されました**"
    # 担当者を変えると割当の通知
    as_owner.post(f"/manage/dev/tasks/{task.pk}/edit/", {
        "title": "通知を試す", "status": "done", "priority": "critical", "category": "infra", "assignee": owner.pk,
    })
    assert len(discord) == 4 and discord[3]["body"]["content"] == "**タスクが割り当てられました**"


def test_Discord_送れなくても保存は成立する(as_owner, project, owner, monkeypatch):
    project.discord_webhook_url = "https://discord.com/api/webhooks/1/abc"
    project.save()

    def broken(req, timeout=0):
        raise OSError("届かない")

    monkeypatch.setattr(notifications, "urlopen", broken)
    r = as_owner.post(f"/manage/dev/{project.pk}/tasks/new/", {
        "title": "届かなくても", "status": "open", "priority": "medium", "category": "feature", "assignee": owner.pk,
    })
    assert r.status_code == 302 and DevTask.objects.filter(title="届かなくても").exists()


# ========== 目安箱 ==========


def test_目安箱_一覧の見出しと列と空の文言(as_owner):
    page = as_owner.get("/manage/dev/meyasubako/").content.decode()
    for s in ["目安箱", "+ ご意見を投稿", "まだ投稿がありません"]:
        assert s in page, s
    for th in ["<th>種別</th>", "<th>対象</th>", "<th>タイトル</th>", "<th>報告者</th>", "<th>緊急度</th>", "<th>状態</th>", "<th>投稿日</th>"]:
        assert th in page, th
    assert 'class="active">目安箱</a>' in page


def test_目安箱_投稿画面の項目と対象の選択肢(as_owner, owner):
    page = as_owner.get("/manage/dev/meyasubako/new/").content.decode()
    for s in ["ご意見・ご要望を投稿", "報告者", "種別", "対象モジュール", "タイトル", "困っていること", "こうなってほしい（任意）",
              "緊急度", "スクリーンショット（任意）", "投稿する", "何が起きましたか？ 何に困っていますか？", "どうなると嬉しいですか？（任意）"]:
        assert s in page, s
    # 入っている人が最初から選ばれている
    assert f'<option value="{owner.pk}" selected>まゆみ</option>' in page
    # 対象はこちらの管理画面の段
    for opt in ["アプリ管理", "公式サイト", "予約管理", "開発", "その他"]:
        assert f">{opt}</option>" in page, opt
    for kem_only in ["現場管理", "日報管理", "原価管理"]:
        assert kem_only not in page
    for kind in ["不具合", "使いにくい", "機能追加"]:
        assert f">{kind}</option>" in page, kind
    assert ">すぐ対応してほしい</option>" in page and ">急がない</option>" in page


def test_目安箱_スクリーンショット付きで投稿し詳細に出る(as_owner, owner, settings):
    r = as_owner.post("/manage/dev/meyasubako/new/", {
        "reporter": owner.pk, "kind": "bug", "module": "reserve", "title": "枠が二重に取れる",
        "problem": "同じ時間に2件", "wish": "止めてほしい", "urgency": "urgent", "screenshot": _png(),
    })
    assert r.status_code == 302 and r["Location"] == "/manage/dev/meyasubako/"
    post = Meyasubako.objects.get()
    assert post.reporter == owner and post.reporter_name == "まゆみ" and post.kind == "bug" and post.module == "reserve"
    assert post.screenshot_url.startswith("https://api.example.com/media/dev/") and post.screenshot_url.endswith(".jpg")
    name = post.screenshot_url.rsplit("/", 1)[1]
    saved = settings.MEDIA_ROOT / "dev" / name
    assert saved.is_file()
    from PIL import Image

    assert max(Image.open(saved).size) <= 1600
    # 配り口
    assert as_owner.get(f"/media/dev/{name}").status_code == 200
    assert as_owner.get("/media/dev/..%2Fnews%2Fx.jpg").status_code == 404

    page = as_owner.get("/manage/dev/meyasubako/").content.decode()
    assert "ご意見を投稿しました。ありがとうございます！" in page
    for s in ['<span class="badge badge-red">不具合</span>', "予約管理", "枠が二重に取れる", "まゆみ",
              '<span class="badge badge-red">すぐ対応</span>', '<span class="badge badge-blue">未対応</span>']:
        assert s in page, s

    detail = as_owner.get(f"/manage/dev/meyasubako/{post.pk}/").content.decode()
    for s in ["困っていること", "同じ時間に2件", "こうなってほしい", "止めてほしい", "スクリーンショット",
              f'<img src="{post.screenshot_url}"', "対応済みにする", "一覧に戻る"]:
        assert s in detail, s


def test_目安箱_スクリーンショット無しでも投稿できる(as_owner, owner, settings):
    r = as_owner.post("/manage/dev/meyasubako/new/", {
        "reporter": owner.pk, "kind": "feature", "module": "app", "title": "一覧に検索", "problem": "探しにくい",
        "urgency": "not_urgent",
    })
    assert r.status_code == 302
    post = Meyasubako.objects.get()
    assert post.screenshot_url == "" and post.urgency == "not_urgent"
    detail = as_owner.get(f"/manage/dev/meyasubako/{post.pk}/").content.decode()
    assert "<h3>スクリーンショット</h3>" not in detail and "<h3>こうなってほしい</h3>" not in detail


def test_目安箱_必須が抜けると戻される(as_owner, owner):
    r = as_owner.post("/manage/dev/meyasubako/new/", {"reporter": owner.pk, "kind": "bug", "module": "app", "title": ""})
    assert r.status_code == 200 and Meyasubako.objects.count() == 0


def test_目安箱_対応済みにして戻す(as_owner, owner):
    post = Meyasubako.objects.create(reporter=owner, reporter_name="まゆみ", kind="other", module="other", title="t", problem="p")
    r = as_owner.post(f"/manage/dev/meyasubako/{post.pk}/resolve/")
    assert r.status_code == 302 and r["Location"] == f"/manage/dev/meyasubako/{post.pk}/"
    post.refresh_from_db()
    assert post.resolved is True
    detail = as_owner.get(f"/manage/dev/meyasubako/{post.pk}/").content.decode()
    assert '<span class="badge badge-green">対応済み</span>' in detail and "未対応に戻す" in detail
    as_owner.post(f"/manage/dev/meyasubako/{post.pk}/resolve/")
    post.refresh_from_db()
    assert post.resolved is False
    # GET では変わらない
    as_owner.get(f"/manage/dev/meyasubako/{post.pk}/resolve/")
    post.refresh_from_db()
    assert post.resolved is False
