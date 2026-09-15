"""公式サイト「✍️ まゆみのつぶやき」。中身は mayumi-site/admin/app.py の sec_blog と
/api/blog-list・blog-post・blog-save・blog-new・blog-preview・blog-del・blog-image と同じ。

記事のデータは blog.json（`repo.部品("blog")` の load()/save()）。**書き方（キー・順序・日付の形）は
旧アプリと同じ**にするため、保存は blog.save をそのまま使い、新しい記事の項目も旧アプリの並びで作る。
写真は旧アプリと同じく「保存」を押す前にファイルだけ先に置き、記事の写真の一覧は保存で確定する。
"""

import datetime
import os
import re

from django.contrib import messages
from django.http import Http404
from django.shortcuts import redirect, render
from django.urls import reverse

from apps.manage.permissions import owner_required

from . import repo, site_images, views

一覧の件数 = 20  # /api/blog-list の per と同じ


def _並べる(posts):
    """新しい順（日付・URLの文字）。旧アプリの一覧と同じ並び。"""
    return sorted(posts, key=lambda p: (p["date"], p["slug"]), reverse=True)


def _ページ送り(page, pages):
    """旧アプリの JS と同じ形: 今の前後2つと最初・最後だけ出し、間は「…」にする。"""
    out = []
    for n in range(1, pages + 1):
        if abs(n - page) > 2 and n not in (1, pages):
            if not out or out[-1] != "…":
                out.append("…")
            continue
        out.append(n)
    return out


def _探す(slug):
    blog = repo.部品("blog")
    d = blog.load()
    for p in d.get("posts", []):
        if p["slug"] == slug:
            return d, p
    return d, None


def _入力から(request, p):
    """画面の入力で記事の写しを作る（まだ保存はしない）。旧アプリの blogSave / blogPreview が
    BCUR に入力を乗せるのと同じ。src など画面に無い項目はそのまま残す。"""
    draft = dict(p)
    draft["date_disp"] = request.POST.get("date_disp", "")
    draft["date"] = request.POST.get("date", "")
    draft["title"] = request.POST.get("title", "")
    draft["text"] = request.POST.get("text", "").replace("\r\n", "\n").replace("\r", "\n")
    draft["images"] = [x for x in request.POST.get("images", "").split("\n") if x.strip()]
    return draft


@owner_required
def hp_blog(request):
    """一覧と検索（sec_blog の上の段と「記事の一覧」）。「＋ 新しく書く」もここで受ける。"""
    if not repo.設定されているか():
        return views._設定なし(request)
    blog = repo.部品("blog")
    if request.method == "POST" and request.POST.get("action") == "new":
        # /api/blog-new と同じ: 今日の日付から URL の文字を決め、先頭に入れて保存し、ページを作り直す
        d = blog.load()
        today = datetime.date.today()
        base = today.strftime("%Y%m%d")
        slug, n = base + "-1", 1
        used = {x["slug"] for x in d["posts"]}
        while slug in used:
            n += 1
            slug = "%s-%d" % (base, n)
        d["posts"].insert(0, {
            "slug": slug, "date": today.isoformat(),
            "date_disp": "%d.%d.%d" % (today.year, today.month, today.day),
            "title": "（新しいつぶやき）",
            "text": "ここに本文を書いてください。\n\n1行が1段落になります。",
            "images": [], "src": ""})
        blog.save(d)
        blog.build()
        messages.success(request, "新しい記事を作りました。書けたら「保存してページを作り直す」を押してください。")
        return redirect(reverse("manage:hp_blog_edit", args=[slug]) + "?new=1")

    word = (request.GET.get("q") or "").strip()
    try:
        page = int(request.GET.get("page") or 1)
    except ValueError:
        page = 1
    posts = _並べる(blog.load().get("posts", []))
    if word:
        w = word.lower()
        posts = [p for p in posts if w in p["title"].lower() or w in p.get("text", "").lower()]
    pages = max(1, (len(posts) + 一覧の件数 - 1) // 一覧の件数)
    page = min(max(1, page), pages)
    chunk = posts[(page - 1) * 一覧の件数: page * 一覧の件数]
    return render(request, "manage/hp_blog.html", {
        "log": views._記録を取る(request),
        "q": word, "page": page, "pages": pages, "pager": _ページ送り(page, pages),
        "total": len(posts),
        "posts": [{"slug": p["slug"], "date_disp": p["date_disp"], "title": p["title"],
                   "n_img": len(p.get("images", []))} for p in chunk],
    })


@owner_required
def hp_blog_edit(request, slug):
    """記事の編集（sec_blog の #blogEdit）。保存・プレビュー・写真の追加・削除。"""
    if not repo.設定されているか():
        return views._設定なし(request)
    blog = repo.部品("blog")
    d, p = _探す(slug)
    if p is None:
        raise Http404()
    core = repo.部品("core")
    見出し = "記事の編集：" + p["title"][:30]
    draft = p
    dirty = False
    preview_url = ""

    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "del":
            # /api/blog-del と同じ
            before = len(d["posts"])
            d["posts"] = [x for x in d["posts"] if x["slug"] != slug]
            blog.save(d)
            blog.build()
            views._記録を出す(request, ["削除しました（%d → %d 記事）。" % (before, len(d["posts"])),
                                  "ページを作り直しました。"], "この記事を削除")
            return redirect("manage:hp_blog")

        draft = _入力から(request, p)
        if 何 == "save":
            # /api/blog-save と同じ: URL の文字は英数字だけにし、重なっていれば止める
            new_slug = (request.POST.get("new_slug") or slug).strip()
            new_slug = re.sub(r"[^A-Za-z0-9_-]", "", new_slug) or slug
            others = [x for x in d["posts"] if x["slug"] != slug]
            if any(x["slug"] == new_slug for x in others):
                messages.error(request, "同じURLの記事がすでにあります。別の文字にしてください。")
                dirty = True
            else:
                p.update(draft)  # 既存の項目の並びを保ったまま入れ替える（blog.json の形を変えない）
                p["slug"] = new_slug
                d["posts"] = others + [p]
                blog.save(d)
                n, pages, _made = blog.build()
                views._記録を出す(request, ["保存しました。",
                                      "ページを作り直しました（記事%d本／一覧%dページ）。" % (n, pages)],
                              "保存してページを作り直す")
                return redirect("manage:hp_blog_edit", new_slug)
        elif 何 == "addimg":
            # /api/blog-image と同じ: 写真はすぐ assets/img/blog/ に置く。記事に付くのは「保存」のとき
            f = request.FILES.get("file")
            if not f:
                messages.error(request, "失敗: ファイルが受け取れませんでした。")
            else:
                try:
                    saved = site_images.写真を足す(f.read(), f.name, "blog", 1200)
                    draft["images"] = draft.get("images", []) + [saved]
                    messages.success(request, "写真を追加しました。「本文に差し込む」で位置を決められます。"
                                              "何も指定しなければ記事の先頭に出ます。")
                except Exception as ex:
                    messages.error(request, "失敗: %s" % ex)
            dirty = True
        elif 何 == "preview":
            # /api/blog-preview と同じ: 入力中の内容で記事ページを admin/_preview に書き、そこを見せる
            draft["slug"] = re.sub(r"[^A-Za-z0-9_-]", "", (request.POST.get("new_slug") or "").strip()) or slug
            try:
                url = blog.write_preview(draft, os.path.join(core.BASE, "_preview"))  # "/draft/blog/<slug>/"
                preview_url = reverse("manage:hp_preview_file", kwargs={"path": url[len("/draft/"):]})
            except Exception as ex:
                messages.error(request, "プレビューを作れませんでした: %s" % ex)
            dirty = True
        new_slug_value = request.POST.get("new_slug", slug)
    else:
        new_slug_value = slug
        if request.GET.get("new"):
            見出し = "新しいつぶやき"

    return render(request, "manage/hp_blog_edit.html", {
        "log": views._記録を取る(request),
        "p": draft, "slug": slug, "new_slug": new_slug_value, "title": 見出し,
        "dirty": dirty, "preview_url": preview_url,
        "images_text": "\n".join(draft.get("images", [])),
    })
