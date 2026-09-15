"""公式サイトの「📢 お知らせ」。中身は mayumi-site/admin/app.py の sec_news と /api/news-* と同じ。

旧アプリは1枚の画面の中で JS が一覧・編集・取り込みを切り替えていた。ここでは
    hp/news/            一覧（トップに出す件数・新しく書く・並べ替え）
    hp/news/<番号>/     1件の編集（写真・本文・保存・プレビュー・削除）
    hp/news/app/        アプリの NEWS から取り込む
の3画面に分けたが、項目・ボタン・文言・保存の結果は同じにしてある。

アプリの NEWS は、旧アプリが GAS（getNews）を叩いて読んでいたところを、
**このサーバーの表（apps.content.models.News）から同じ形（apps.gasapi.views._お知らせ）で読む。**
"""

import datetime
import os
import re
import urllib.request

from django.contrib import messages
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_http_methods

from apps.manage.permissions import owner_required

from . import repo, site_images, views

# 写真1枚を待つ上限（app.py の APP_IMG_TIMEOUT と同じ）。
# Googleドライブが応答しないものが混ざっていても全体が止まらないよう短くしてある。
APP_IMG_TIMEOUT = 15


# ------------------------------------------------------------------ 小さな道具（app.py から写した）

def _日付表示(iso):
    """JS の dispDate と同じ。"2026-09-02" → "2026.9.2"。"""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", (iso or "").strip())
    return "%d.%d.%d" % (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else ""


def _次のid(news):
    """next_news_id と同じ。空いている番号の一番小さいものを使う。"""
    used = {n.get("id") for n in news}
    i = 1
    while "news-%02d" % i in used:
        i += 1
    return "news-%02d" % i


def _本文にする(n):
    """編集欄に出す本文。

    旧アプリは text だけを見ていたので、昔の形（blocks だけ）のお知らせを開くと
    本文が空で出て、そのまま保存すると本文が消えた。ここでは blocks を同じ書き方
    （# 見出し・- 箇条書き・{img:…}）の文章に直して出す。囲み枠（note）は
    ふつうの段落になる（旧アプリで保存し直したときと同じ）。
    """
    if (n.get("text") or "").strip():
        return n["text"]
    parts = []
    for b in n.get("blocks") or []:
        t, v = b.get("t"), b.get("v")
        if t == "p":
            parts.append(str(v or ""))
        elif t == "h3":
            parts.append("# " + str(v or ""))
        elif t == "list":
            parts.append("\n".join("- " + str(x) for x in (v or [])))
        elif t == "note":
            parts.append("\n".join(str(x) for x in (v or [])))
        elif t == "img":
            parts.append("{img:%s}" % v)
    return "\n\n".join(p for p in parts if p)


def _ページを作り直す(d, log, 件数つき=False):
    """news-save / news-del / news-move の後半と同じ。お知らせのページ・sitemap・トップの3件を作り直す。"""
    news = repo.部品("news")
    core = repo.部品("core")
    try:
        cnt = news.build()
        news.update_sitemap()
        core.apply_blocks(d)          # トップページの3件も更新
        log.append("ページを作り直しました（%d件）。" % cnt if 件数つき else "ページを作り直しました。")
    except Exception as ex:
        log.append("ページの作り直しに失敗: %s" % ex)
    return log


def _選ぶ(f, category):
    """JS の pickBy と同じ。空は全部、__none__ は種別なし、それ以外は一致。"""
    c = (category or "").strip()
    if not f:
        return True
    if f == "__none__":
        return c == ""
    return c == f


# ------------------------------------------------------------------ 一覧

@owner_required
def hp_news(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    d = core.load()
    d.setdefault("news", [])
    cat = (request.POST.get("cat") if request.method == "POST" else request.GET.get("cat")) or ""

    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "new":
            # /api/news-new と同じ。今日の日付で先頭に足し、すぐ編集欄を開く
            today = datetime.date.today()
            item = {
                "id": _次のid(d["news"]),
                "date": today.isoformat(),
                "date_disp": "%d.%d.%d" % (today.year, today.month, today.day),
                "title": "（新しいお知らせ）",
                "category": "",
                "blocks": [{"t": "p", "v": "ここに本文を書いてください。"}],
                "images": [],
            }
            d["news"].insert(0, item)
            core.save(d)
            messages.success(request, "新しいお知らせを作りました。書けたら「保存してページを作り直す」を押してください。")
            return redirect("manage:hp_news_edit", index=0)
        if 何 == "move":
            # /api/news-move と同じ。隣と入れ替えてページを作り直す
            try:
                a, b = int(request.POST.get("from", "")), int(request.POST.get("to", ""))
            except ValueError:
                a = b = -1
            if 0 <= a < len(d["news"]) and 0 <= b < len(d["news"]):
                d["news"][a], d["news"][b] = d["news"][b], d["news"][a]
                core.save(d)
                _ページを作り直す(d, [])
                messages.success(request, "並べ替えました。")
        elif 何 == "topcount":
            # 旧アプリでは「手元に反映」のときに他の項目と一緒に保存していたもの。ここでは単独で保存する
            try:
                d["news_top_count"] = int(request.POST.get("news_top_count", "")) or 3
            except ValueError:
                d["news_top_count"] = 3
            core.save(d)
            messages.success(request, "トップに出す件数を保存しました。")
        url = reverse("manage:hp_news")
        return redirect(url + ("?cat=" + cat if cat else ""))

    rows = [(i, n) for i, n in enumerate(d["news"]) if _選ぶ(cat, n.get("category"))]
    return render(request, "manage/hp_news.html", {
        "log": views._記録を取る(request),
        "total": len(d["news"]), "top_count": d.get("news_top_count", 3),
        "categories": [o.get("name", "") for o in core.options(d, "news_categories")],
        "cat": cat, "rows": rows, "last": len(d["news"]) - 1,
    })


# ------------------------------------------------------------------ 編集

def _入力を読む(request, n):
    """JS の nGrab と同じ。入力欄の内容を1件分にまとめる。段落ごとの箱（blocks）はもう使わない。"""
    item = dict(n)
    item["date"] = (request.POST.get("date") or "").strip()
    item["date_disp"] = _日付表示(item["date"]) or n.get("date_disp", "")
    item["title"] = request.POST.get("title") or ""
    item["category"] = request.POST.get("category") or ""
    item["text"] = (request.POST.get("text") or "").replace("\r\n", "\n").replace("\r", "\n")
    item["images"] = [x for x in request.POST.getlist("images") if x]
    item.pop("blocks", None)
    return item


def _編集画面(request, d, index, item, dirty=False, preview=None, label=None):
    core = repo.部品("core")
    return render(request, "manage/hp_news_edit.html", {
        "log": views._記録を取る(request),
        "index": index, "item": item, "text": item.get("text") or "",
        "label": label or ("お知らせの編集：" + (item.get("title") or "")[:30]),
        "categories": [o.get("name", "") for o in core.options(d, "news_categories")],
        "dirty": dirty, "preview": preview,
        "img_base": reverse("manage:hp_site_file", kwargs={"path": "assets/img/news/"}),
    })


@owner_required
def hp_news_edit(request, index):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    news = repo.部品("news")
    d = core.load()
    d.setdefault("news", [])
    if not (0 <= index < len(d["news"])):
        messages.error(request, "その位置のお知らせがありません。")
        return redirect("manage:hp_news")
    n = d["news"][index]

    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "delete":
            # /api/news-del と同じ
            gone = d["news"].pop(index)
            core.save(d)
            log = ["「%s」を削除しました。" % gone.get("title", "")]
            _ページを作り直す(d, log)
            views._記録を出す(request, log, "このお知らせを削除")
            return redirect("manage:hp_news")
        item = _入力を読む(request, n)
        if 何 == "save":
            # /api/news-save と同じ
            d["news"][index] = item
            core.save(d)
            log = ["保存しました。"]
            _ページを作り直す(d, log, 件数つき=True)
            views._記録を出す(request, log, "保存してページを作り直す")
            return redirect("manage:hp_news_edit", index=index)
        if 何 == "add_image":
            # /api/news-image と同じ。写真だけ保存し、お知らせ自体はまだ保存しない（未保存のまま編集欄に戻す）
            f = request.FILES.get("file")
            if not f:
                messages.error(request, "ファイルが受け取れませんでした。")
            else:
                try:
                    item["images"].append(site_images.写真を足す(f.read(), f.name, "news", 1200))
                    messages.success(request, "写真を追加しました。「本文に差し込む」で位置を決められます。")
                except Exception as ex:
                    messages.error(request, "失敗: %s" % ex)
            return _編集画面(request, d, index, item, dirty=True)
        if 何 == "preview":
            # /api/news-preview と同じ。いま入力している内容で1ページだけ書き出す（保存はしない）
            preview = None
            try:
                news.write_preview(item, os.path.join(core.BASE, "_preview"))
                preview = reverse("manage:hp_preview_file",
                                  kwargs={"path": "news/%s/" % (item.get("id") or "preview")})
            except Exception as ex:
                messages.error(request, "プレビューを作れませんでした: %s" % ex)
            return _編集画面(request, d, index, item, dirty=True, preview=preview)
        return redirect("manage:hp_news_edit", index=index)

    item = dict(n)
    item["text"] = _本文にする(n)
    label = "新しいお知らせ" if n.get("title") == "（新しいお知らせ）" else None
    return _編集画面(request, d, index, item, label=label)


# ------------------------------------------------------------------ アプリの NEWS から取り込む

def _合図(date, title):
    """app_news_key と同じ。同じお知らせかどうかは日付と題名の組で見る（行番号は消すとずれる）。"""
    return "%s|%s" % ((date or "").strip(), " ".join((title or "").split()))


def _アプリの日付(raw):
    """app_news_date と同じ。"2026-09-07T00:00:00+09:00" も "2026.09.07" も "2026-09-07" にする。"""
    t = str(raw or "").strip()
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", t)
    if not m:
        return ""
    y, mo, da = (int(x) for x in m.groups())
    try:
        datetime.date(y, mo, da)
    except ValueError:
        return ""
    return "%04d-%02d-%02d" % (y, mo, da)


def _アプリの段落(body):
    """app_news_blocks と同じ。空行が段落の区切り。段落の中の改行はそのまま残す。"""
    t = str(body or "").replace("\r\n", "\n").replace("\r", "\n")
    paras = [p.strip("\n").strip() for p in re.split(r"\n\s*\n", t)]
    return [{"t": "p", "v": p} for p in paras if p]


def _アプリのNEWS(d):
    """fetch_app_news と同じ。GAS の代わりに、このサーバーの表から同じ形（_お知らせ）で読む。"""
    from apps.gasapi.views import _お知らせ

    payload = _お知らせ()
    already = {_合図(n.get("date"), n.get("title")) for n in d.get("news", [])}
    out = []
    for n in payload.get("news", []):
        date = _アプリの日付(n.get("date"))
        title = " ".join(str(n.get("title") or "").split())
        if not title:
            continue
        imgs = [u for u in (n.get("imageUrls") or []) if str(u).startswith("http")]
        out.append({
            "key": _合図(date, title),
            "date": date,
            "title": title,
            "category": str(n.get("category") or "").strip(),
            "body": str(n.get("body") or ""),
            "images": imgs,
            "done": _合図(date, title) in already,
        })
    out.sort(key=lambda x: (x["date"], x["title"]), reverse=True)
    return out


def _写真を取る(url):
    """写真1枚を URL から読む（試験では差し替える）。"""
    rq = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(rq, timeout=APP_IMG_TIMEOUT) as r:
        return r.read()


def _取り込む(keys):
    """import_app_news と同じ。選ばれたNEWSをサイトのお知らせにする。写真も持ってくる。"""
    core = repo.部品("core")
    d = core.load()
    d.setdefault("news", [])
    want = set(keys or [])
    items = [x for x in _アプリのNEWS(d) if x["key"] in want and not x["done"]]
    if not items:
        return ["取り込むものがありませんでした。"], 0

    log = []
    added = 0
    # 古いものから入れて、最後に新しい順へ並べ直す
    for x in sorted(items, key=lambda a: (a["date"], a["title"])):
        names = []
        for u in x["images"]:
            try:
                names.append(site_images.写真を足す(_写真を取る(u), "app_" + x["date"].replace("-", ""), "news", 1200))
            except Exception as ex:
                # 写真が取れなくても本文は入れる。あとから手で足せる。
                log.append("　写真を取れませんでした（%s）: %s" % (x["title"][:18], ex))
        y, mo, da = (int(v) for v in x["date"].split("-")) if x["date"] else (0, 0, 0)
        d["news"].append({
            "id": _次のid(d["news"]),
            "date": x["date"],
            "date_disp": ("%d.%d.%d" % (y, mo, da)) if x["date"] else "",
            "title": x["title"],
            "category": x["category"],
            "blocks": _アプリの段落(x["body"]),
            "images": names,
        })
        added += 1
        log.append("取り込みました：%s %s" % (x["date"], x["title"][:30]))

    d["news"].sort(key=lambda n: (n.get("date") or ""), reverse=True)
    core.save(d)
    log.insert(0, "%d件を取り込みました。" % added)
    return log, added


@owner_required
@require_http_methods(["GET", "POST"])
def hp_news_app(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")

    if request.method == "POST":
        keys = [k for k in request.POST.getlist("keys") if k]
        if not keys:
            messages.error(request, "取り込むものが選ばれていません。")
            return redirect("manage:hp_news_app")
        try:
            log, added = _取り込む(keys)
        except Exception as ex:
            log, added = ["取り込みに失敗しました: %s" % ex], 0
        if added:
            # /api/app-news-import と同じ。取り込めたらページを作り直す
            d = core.load()
            try:
                repo.部品("news").build()
                repo.部品("news").update_sitemap()
                core.apply_blocks(d)      # トップページの3件も入れ替える
                log.append("お知らせのページを作り直しました。")
            except Exception as ex:
                log.append("ページの作り直しに失敗: %s" % ex)
        views._記録を出す(request, log, "📥 アプリのNEWSから取り込む")
        return redirect("manage:hp_news")

    cat = request.GET.get("cat") or ""
    only_new = request.GET.get("only_new", "1") != "0"
    error = None
    try:
        items = _アプリのNEWS(core.load())
    except Exception as ex:
        items, error = [], "アプリのNEWSを読めませんでした: %s" % ex
    rows = [x for x in items if (not cat or x["category"] == cat) and (not only_new or not x["done"])]
    yet = sum(1 for x in items if not x["done"])
    if rows:
        empty = ""
    elif only_new and yet == 0:
        empty = "アプリのNEWSは、すべて公式サイトに出してあります。"
    else:
        empty = "見せるものがありません。種別のしぼり込みを変えてみてください。"
    return render(request, "manage/hp_news_app.html", {
        "error": error, "rows": rows, "cat": cat, "only_new": only_new,
        "cats": sorted({x["category"] for x in items if x["category"]}),
        "count_all": len(items), "count_yet": yet, "count_shown": len(rows), "empty": empty,
    })
