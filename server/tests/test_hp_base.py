"""公式サイト（apps/hp）: 保存の記録・公開する・プレビュー。中身は mayumi-site/admin/app.py と同じ。"""

import json
import os
import subprocess

import pytest

pytestmark = pytest.mark.django_db


def _git(cwd, *args):
    return subprocess.run(["git", *args], cwd=cwd, capture_output=True, text=True).stdout.strip()


def test_設定が無ければ案内だけ出て落ちない(as_owner, settings):
    settings.SITE_REPO_DIR = ""
    for u in ("/manage/hp/git/", "/manage/hp/publish/", "/manage/hp/preview/"):
        r = as_owner.get(u)
        assert r.status_code == 200 and "まだ設定されていません" in r.content.decode(), u


def test_スタッフは入れない(client, staff, site_repo):
    client.force_login(staff)
    assert client.get("/manage/hp/git/").status_code in (302, 403)


def test_保存の記録_状態と下書き保存(as_owner, site_repo):
    page = as_owner.get("/manage/hp/git/").content.decode()
    assert "保存の記録" in page and "draft" in page and "保存していない変更はありません" in page
    # content.json を変えると「まだ保存していない変更」が出る
    p = site_repo / "admin" / "content.json"
    d = json.loads(p.read_text(encoding="utf-8"))
    d["news_top_count"] = int(d.get("news_top_count") or 3) + 1
    p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    page = as_owner.get("/manage/hp/git/").content.decode()
    # gitops.run が出力を strip するので porcelain の先頭1文字が落ちる（旧アプリからの癖）。名前の末尾で見る
    assert "まだ保存していない変更が 1 件あります" in page and "content.json" in page
    # 下書きとして保存する → draft に commit され origin へ送られる
    r = as_owner.post("/manage/hp/git/", {"action": "commit_push", "message": "試験の保存"})
    assert r.status_code == 302
    page = as_owner.get("/manage/hp/git/").content.decode()
    assert "試験の保存" in page and "保存していない変更はありません" in page
    assert _git(site_repo, "branch", "--show-current") == "draft"
    assert "試験の保存" in _git(site_repo, "log", "-1", "--pretty=%s", "origin/draft")


def test_受け取る_向こうが進んでいなければその旨(as_owner, site_repo):
    as_owner.post("/manage/hp/git/", {"action": "pull"})
    page = as_owner.get("/manage/hp/git/").content.decode()
    assert "受け取るものはありません" in page


def test_公開する_確認用を通してから公開できる(as_owner, site_repo):
    page = as_owner.get("/manage/hp/publish/").content.decode()
    assert "サイトに反映（確認用）" in page and "サイトに反映（公開）" in page and 'value="go" disabled' in page
    assert "いま管理している場所" in page and "目印の名前" in page
    # 確認用を押していないのに公開は押せない
    r = as_owner.post("/manage/hp/publish/", {"action": "go"}, follow=True)
    assert "先に「サイトに反映（確認用）」を押して" in r.content.decode()
    # 変更を1つ入れて、確認用 → 公開
    p = site_repo / "admin" / "content.json"
    d = json.loads(p.read_text(encoding="utf-8"))
    d["news_top_count"] = int(d.get("news_top_count") or 3) + 1
    p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    r = as_owner.post("/manage/hp/publish/", {"action": "prepare", "message": "試験の公開"}, follow=True)
    page = r.content.decode()
    assert "バックアップ:" in page and "GitHubに接続できます" in page, page[:3000]
    assert "公開できます" in page and 'value="go" disabled' not in page
    r = as_owner.post("/manage/hp/publish/", {"action": "go"}, follow=True)
    page = r.content.decode()
    assert "公開しました" in page and "下書き（draft）に戻りました" in page
    assert _git(site_repo, "branch", "--show-current") == "draft"
    assert _git(site_repo, "rev-parse", "origin/main") == _git(site_repo, "rev-parse", "draft")
    # 控えができ、「元に戻す」に並ぶ
    page = as_owner.get("/manage/hp/publish/").content.decode()
    assert "この時点に戻す" in page and 'value="go" disabled' in page


def test_元に戻す(as_owner, site_repo):
    as_owner.post("/manage/hp/publish/", {"action": "local"})
    backups = sorted(os.listdir(site_repo / "admin" / "backups"))
    assert backups
    index = site_repo / "index.html"
    index.write_text(index.read_text(encoding="utf-8") + "<!-- 壊した -->", encoding="utf-8")
    r = as_owner.post("/manage/hp/publish/", {"action": "restore", "name": backups[-1]}, follow=True)
    assert "の時点に戻しました" in r.content.decode()
    assert "<!-- 壊した -->" not in index.read_text(encoding="utf-8")
    # 変な名前は受け付けない
    r = as_owner.post("/manage/hp/publish/", {"action": "restore", "name": "../admin"}, follow=True)
    assert "見つかりませんでした" in r.content.decode()


def test_プレビュー_編集中の内容で作り直して見られる(as_owner, site_repo):
    page = as_owner.get("/manage/hp/preview/").content.decode()
    assert "プレビュー" in page and "トップ" in page and "編集中の内容で見る" in page
    r = as_owner.post("/manage/hp/preview/", follow=True)
    assert "作り直しました" in r.content.decode()
    assert (site_repo / "admin" / "_preview" / "index.html").is_file()
    r = as_owner.get("/manage/hp/preview/files/")
    assert r.status_code == 200 and b"<html" in b"".join(r.streaming_content).lower()
    r = as_owner.get("/manage/hp/preview/files/news/")
    assert r.status_code == 200
    # admin/ や docs/ は配らない
    assert as_owner.get("/manage/hp/site/admin/content.json").status_code == 404
    assert as_owner.get("/manage/hp/site/../admin/content.json").status_code in (404, 400)
    assert as_owner.get("/manage/hp/site/").status_code == 200


def test_枠に出すものは同じ出どころなら許す(as_owner, site_repo):
    """既定の X-Frame-Options: DENY のままだと、プレビューの枠が「接続が拒否されました」になる（2026-09-16 に実地で発生）。"""
    as_owner.post("/manage/hp/preview/")
    for u in ("/manage/hp/preview/files/", "/manage/hp/site/", "/manage/hp/layout/pv/access/"):
        r = as_owner.get(u)
        assert r.status_code == 200, u
        assert r.get("X-Frame-Options", "").upper() == "SAMEORIGIN", (u, r.get("X-Frame-Options"))
    # 管理画面そのものは今までどおり枠に入れさせない
    assert as_owner.get("/manage/hp/git/").get("X-Frame-Options", "").upper() == "DENY"
