"""公式サイトの画面（公開する・保存の記録・プレビュー）。中身は mayumi-site/admin/app.py と同じ。"""

import mimetypes
import os

from django.contrib import messages
from django.http import FileResponse, Http404, HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.clickjacking import xframe_options_sameorigin
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import actions, repo

# プレビューで選べるページ（app.py の PAGES と同じ）
PAGES = [
    ("/", "トップ"), ("/outpatient/", "母乳外来"), ("/care02/", "産後ケア"),
    ("/itothermy/", "イトオテルミー"), ("/acupuncturist/", "鍼灸・整体"),
    ("/reception/", "料金・診察時間"), ("/staff/", "スタッフ紹介"),
    ("/classroom/", "各種お教室"), ("/access/", "アクセス"),
    ("/news/", "お知らせ"), ("/404.html", "404ページ"),
]
DEVICES = [("スマホ", 390), ("タブレット", 820), ("パソコン", 0)]


def _設定なし(request):
    return render(request, "manage/hp_unconfigured.html", status=200)


def _記録を出す(request, log, 見出し=""):
    """処理の結果（複数行）を次の画面に渡す。"""
    request.session["hp_log"] = {"title": 見出し, "lines": [str(x) for x in log]}


def _記録を取る(request):
    return request.session.pop("hp_log", None)


@owner_required
def hp_git(request):
    if not repo.設定されているか():
        return _設定なし(request)
    gitops = repo.部品("gitops")
    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "pull":
            _記録を出す(request, gitops.pull(), "📥 GitHubから受け取る")
        elif 何 in ("commit_push", "commit"):
            _記録を出す(request, gitops.commit_push(request.POST.get("message", "").strip(), push=(何 == "commit_push")),
                    "下書きとして保存する" if 何 == "commit_push" else "記録だけ（送信しない）")
        return redirect("manage:hp_git")
    状態 = gitops.status()
    return render(request, "manage/hp_git.html", {"st": 状態, "log": _記録を取る(request), "repo_dir": repo.場所()})


@owner_required
def hp_publish(request):
    if not repo.設定されているか():
        return _設定なし(request)
    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "prepare":
            ok, log = actions.公開の下ごしらえ(request.POST.get("message", "").strip())
            request.session["hp_publish_ready"] = bool(ok)
            _記録を出す(request, log, "サイトに反映（確認用）" + ("　→ 公開できます" if ok else "　→ ここで止まりました"))
        elif 何 == "go":
            if not request.session.get("hp_publish_ready"):
                messages.error(request, "先に「サイトに反映（確認用）」を押して、問題がないことを確かめてください。")
                return redirect("manage:hp_publish")
            ok, log = actions.公開する()
            request.session["hp_publish_ready"] = False
            _記録を出す(request, log, "サイトに反映（公開）")
        elif 何 == "local":
            core = repo.部品("core")
            _記録を出す(request, actions.反映する(core.load()), "手元に反映")
        elif 何 == "restore":
            _記録を出す(request, actions.控えに戻す(request.POST.get("name", "")), "↩ 元に戻す")
            request.session["hp_publish_ready"] = False
        return redirect("manage:hp_publish")
    return render(request, "manage/hp_publish.html", {
        "log": _記録を取る(request), "ready": bool(request.session.get("hp_publish_ready")),
        "backups": actions.控えの一覧(), "markers": actions.目印の状態(),
    })


@owner_required
def hp_preview(request):
    if not repo.設定されているか():
        return _設定なし(request)
    if request.method == "POST":
        # 「編集中の内容で見る」: いまの下書き（content.json）でプレビュー用の HTML を作り直す
        core = repo.部品("core")
        try:
            core.write_preview(core.load(), os.path.join(core.BASE, "_preview"))
            messages.success(request, "編集中の内容でプレビューを作り直しました。")
        except Exception as ex:
            messages.error(request, "プレビューを作れませんでした: %s" % ex)
        return redirect(request.POST.get("next") or "manage:hp_preview")
    page = request.GET.get("page") or "/"
    if page not in {p for p, _ in PAGES}:
        page = "/"
    draft = request.GET.get("draft", "1") != "0"
    width = request.GET.get("w", "390")
    return render(request, "manage/hp_preview.html", {
        "pages": PAGES, "page": page, "draft": draft, "devices": DEVICES, "width": width,
    })


def _配る(基点: str, パス: str):
    """基点の中のファイルだけを返す（admin/ docs/ .git は返さない）。"""
    パス = (パス or "").lstrip("/")
    if パス == "" or パス.endswith("/"):
        パス += "index.html"
    先頭 = パス.split("/")[0]
    if 先頭 in ("admin", "docs", ".git") or ".." in パス.split("/"):
        raise Http404()
    実体 = os.path.abspath(os.path.join(基点, パス))
    if not 実体.startswith(os.path.abspath(基点) + os.sep) or not os.path.isfile(実体):
        raise Http404()
    種類, _ = mimetypes.guess_type(実体)
    return FileResponse(open(実体, "rb"), content_type=種類 or "application/octet-stream")


@owner_required
@xframe_options_sameorigin  # 管理画面の枠（iframe）に出すため。既定の DENY だと Chrome が「接続が拒否されました」と出す
def hp_preview_file(request, path=""):
    """編集中の内容で作ったプレビュー（admin/_preview）。無いページは手元のサイトの方を返す。"""
    core = repo.部品("core")
    pv = os.path.join(core.BASE, "_preview")
    try:
        return _配る(pv, path)
    except Http404:
        return _配る(core.SITE, path)


@owner_required
@xframe_options_sameorigin  # 管理画面の枠（iframe）に出すため。既定の DENY だと Chrome が「接続が拒否されました」と出す
def hp_site_file(request, path=""):
    """手元のサイト（下書きの作業ツリー）。「手元に反映」したものが見える。"""
    core = repo.部品("core")
    return _配る(core.SITE, path)
