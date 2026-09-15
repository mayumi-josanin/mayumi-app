"""公式サイトの「🎒 各種お教室」。中身は mayumi-site/admin/app.py の sec_classes と同じ。

旧アプリは編集中のお教室を画面（JavaScript）に持ち、「保存してページを作り直す」で送っていた。
ここでは **編集中の中身をフォームに持たせ**、写真を足す・段落を足す・回を足すなどの操作は
フォームごと送って読み直し、同じ画面に返す（content.json にはまだ書かない）。
content.json に書くのは「保存してページを作り直す」「これを削除」「↑↓」「＋ 新しいお教室」だけ。
"""

import os
import re

from django.contrib import messages
from django.http import Http404
from django.shortcuts import redirect, render
from django.urls import reverse

from apps.manage.permissions import owner_required

from . import actions, repo, site_images, views

# 1回ずつの分類（classroom.REPORT_KINDS と同じ色）
_KIND_COLOR = {"開催予定": "#e08a3c", "過去の開催": "#7a9c5e"}


# ------------------------------------------------------------------ 読み書き

def _名前(c):
    return (c.get("category") or c.get("title") or "").strip()


def _新しいお教室(d):
    """app.py の /api/class-new と同じ。空のお教室を先頭に足す。"""
    d.setdefault("classroom", [])
    used = {x.get("id") for x in d["classroom"]}
    k = 1
    while "class-%02d" % k in used:
        k += 1
    item = {
        "id": "class-%02d" % k, "category": "", "kind": "", "title": "",
        "date_disp": "", "open_date": "", "open_disp": "", "date": "",
        "show": True, "lead": "", "images": [],
        "blocks": [{"t": "p", "v": "どんなお教室かを書いてください。"}],
        "reports": [],
    }
    d["classroom"].insert(0, item)
    return item


def _ページを作り直す(d, log=None):
    """保存のあとにページを作り直す（app.py の class-save と同じ順）。"""
    classroom = repo.部品("classroom")
    core = repo.部品("core")
    cnt = classroom.build()
    classroom.update_sitemap()
    core.apply_blocks(d)          # お教室ページの一覧も更新
    if log is not None:
        log.append("ページを作り直しました（%d件を掲載）。" % cnt)
    return cnt


def _保存の形にする(item):
    """app.py の /api/class-save がしている整え方。"""
    # 見出し・パンくず・ページ名にはお教室の名称を使う
    item["title"] = (item.get("category") or "").strip() or item.get("title", "")
    # 「過去の開催」に、ページのURLになる番号を付ける。
    # 一度付けたら変えない（並べ替えてもURLが変わらないようにするため）。
    reps = item.get("reports") or []
    used = {(r.get("id") or "").strip() for r in reps if (r.get("id") or "").strip()}
    k = 1
    for r in reps:
        if not (r.get("id") or "").strip():
            while ("r%02d" % k) in used:
                k += 1
            r["id"] = "r%02d" % k
            used.add(r["id"])
    # 並べ替え用の日付は、いちばん新しい「過去の開催」の日付から取る
    dates = [r.get("date") for r in reps if (r.get("date") or "").strip()]
    item["date"] = max(dates) if dates else (item.get("date") or "")
    return item


def _古くなったお教室(d):
    """app.py の stale_classes と同じ。開催日を過ぎたのに、サイトのHTMLがまだ古い分類のままのお教室。

    分類は日付から決まるが、サイトは静的HTMLなので「手元に反映」するまで書き換わらない。
    押し忘れると終わった教室が「開催予定」のまま出続けるので、ここで気づけるようにする。
    """
    core = repo.部品("core")
    classroom = repo.部品("classroom")
    out = []
    for c in d.get("classroom", []):
        cid = (c.get("id") or "").strip()
        page = os.path.join(core.SITE, "classroom", cid, "index.html")
        if not cid or not os.path.isfile(page):
            continue
        try:
            html = open(page, encoding="utf-8").read()
        except OSError:
            continue
        # 数えるのは repList の中のカードだけ。その手前のしぼり込みボタンにも data-rkind が付いている
        body = html.split('id="repList"', 1)
        if len(body) < 2:
            continue
        built = re.findall(r'data-rkind="([^"]*)"', body[1])
        now = [classroom.report_kind(r) for r in classroom.visible_reports(c)]
        if built and built != now:
            out.append((cid, _名前(c) or cid))
    return out


# ------------------------------------------------------------ フォーム ⇄ お教室

def _フォームから読む(P, 元=None):
    """編集画面の入力を、content.json に入る形のお教室に戻す（旧アプリの crGrab と同じ）。"""
    item = dict(元 or {})
    item["id"] = (P.get("id") or item.get("id") or "").strip()
    item["category"] = P.get("category", "")
    item["kind"] = P.get("kind", "")
    item["title"] = item["category"]
    item["date_disp"] = P.get("date_disp", "")
    item["open_date"] = P.get("open_date", "")
    item["open_disp"] = P.get("open_disp", "")
    item["lead"] = P.get("lead", "")
    item["show"] = P.get("show", "1") == "1"
    item["date"] = P.get("date", item.get("date", ""))
    item["images"] = [f for f in P.getlist("images") if f]
    blocks = []
    i = 0
    while ("blk_t_%d" % i) in P:
        t = P.get("blk_t_%d" % i)
        v = P.get("blk_v_%d" % i, "")
        if t in ("p", "h3", "img"):
            blocks.append({"t": t, "v": v})
        else:
            blocks.append({"t": t, "v": [x for x in v.replace("\r\n", "\n").split("\n") if x.strip() != ""]})
        i += 1
    item["blocks"] = blocks
    reps = []
    i = 0
    while ("rep_title_%d" % i) in P:
        r = {
            "title": P.get("rep_title_%d" % i, ""), "date": P.get("rep_date_%d" % i, ""),
            "date_disp": P.get("rep_disp_%d" % i, ""), "kind": P.get("rep_kind_%d" % i, ""),
            "text": P.get("rep_text_%d" % i, "").replace("\r\n", "\n"),
            "images": [f for f in P.getlist("rep_img_%d" % i) if f],
        }
        rid = (P.get("rep_id_%d" % i) or "").strip()
        if rid:
            r["id"] = rid
        reps.append(r)
        i += 1
    item["reports"] = reps
    return item


def _分類の見せ方(r):
    """旧アプリの crKindBadge と同じ。分類は日付だけで決まる。"""
    classroom = repo.部品("classroom")
    kind = classroom.report_kind(r)
    if classroom.report_date(r) is None:
        why = "日付が入っていないため"
    else:
        why = "これから開催" if kind == "開催予定" else "開催日を過ぎたため"
    return kind, _KIND_COLOR[kind], why


def _画面用(item):
    """テンプレートで扱いやすい形にする（段落の名前・回の番号・分類の色）。"""
    NAME = {"p": "段落", "h3": "小見出し", "list": "チェックリスト", "note": "囲み枠", "img": "写真"}
    blocks = []
    for bi, b in enumerate(item.get("blocks") or []):
        v = b.get("v", "")
        blocks.append({
            "i": bi, "t": b.get("t", ""), "name": NAME.get(b.get("t"), b.get("t")),
            "val": v if isinstance(v, str) else "\n".join(v),
            "hint": "" if b.get("t") in ("p", "h3") else "1行が1項目になります。",
        })
    reps = []
    n = len(item.get("reports") or [])
    for ri, r in enumerate(item.get("reports") or []):
        kind, color, why = _分類の見せ方(r)
        reps.append({"i": ri, "no": n - ri, "r": r, "kind": kind, "color": color, "why": why,
                     "text": r.get("text", ""), "images": r.get("images") or []})
    return {"blocks": blocks, "reports": reps}


def _選択肢(d):
    core = repo.部品("core")
    return ([o.get("name", "") for o in core.options(d, "class_categories")],
            [o.get("name", "") for o in core.options(d, "class_kinds")])


# ------------------------------------------------------------------- 一覧

@owner_required
def hp_classes(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    if request.method == "POST":
        何 = request.POST.get("action", "")
        d = core.load()
        if 何 == "new":
            _新しいお教室(d)
            core.save(d)
            messages.success(request, "新しく作りました。書けたら「保存してページを作り直す」を押してください。")
            return redirect(reverse("manage:hp_class_edit", args=[0]) + "?new=1")
        if 何 == "move":
            cs = d.setdefault("classroom", [])
            try:
                a = int(request.POST.get("i", ""))
            except ValueError:
                a = -1
            b = a + (1 if request.POST.get("dir") == "down" else -1)
            if 0 <= a < len(cs) and 0 <= b < len(cs):
                cs[a], cs[b] = cs[b], cs[a]
                core.save(d)
                _ページを作り直す(d)
                messages.success(request, "並べ替えました。")
        if 何 == "local":
            views._記録を出す(request, actions.反映する(d), "手元に反映")
        return redirect(request.POST.get("next") or "manage:hp_classes")

    d = core.load()
    cs = d.get("classroom", [])
    cats, kinds = _選択肢(d)
    cat = request.GET.get("cat", "")
    kind = request.GET.get("kind", "")

    def hit(v, f):
        c = (v or "").strip()
        if not f:
            return True
        if f == "__none__":
            return c == ""
        return c == f

    rows = [{"i": i, "c": c, "name": _名前(c), "n_img": len(c.get("images") or []),
             "n_rep": len(c.get("reports") or []), "hidden": c.get("show") is False}
            for i, c in enumerate(cs) if hit(c.get("category"), cat) and hit(c.get("kind"), kind)]
    count = ("%d件 / 全%d件" % (len(rows), len(cs))) if (cat or kind) else ("全%d件" % len(cs))
    stale = _古くなったお教室(d)
    return render(request, "manage/hp_classes.html", {
        "log": views._記録を取る(request), "total": len(cs), "rows": rows, "count": count,
        "cats": cats, "kinds": kinds, "cat": cat, "kind": kind,
        "stale": "、".join(nm for _i, nm in stale),
    })


# ------------------------------------------------------------------- 編集

@owner_required
def hp_class_edit(request, i):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    classroom = repo.部品("classroom")
    d = core.load()
    cs = d.setdefault("classroom", [])
    if not (0 <= i < len(cs)):
        raise Http404("その位置のお教室がありません。")
    cats, kinds = _選択肢(d)
    preview_url = ""
    dirty = False

    if request.method == "POST":
        P = request.POST
        何 = P.get("action", "")
        item = _フォームから読む(P, cs[i])
        dirty = True

        if 何 == "save":
            cs[i] = _保存の形にする(item)
            core.save(d)
            log = ["保存しました。"]
            try:
                _ページを作り直す(d, log)
            except Exception as ex:
                log.append("ページの作り直しに失敗: %s" % ex)
            views._記録を出す(request, log, "保存してページを作り直す")
            return redirect("manage:hp_class_edit", i)

        if 何 == "del":
            gone = cs.pop(i)
            core.save(d)
            _ページを作り直す(d)
            messages.success(request, "「%s」を削除しました。" % (_名前(gone) or "名前なしのお教室"))
            return redirect("manage:hp_classes")

        # ---- ここから下は content.json に書かない。入力を持ったまま同じ画面に返す ----
        if 何 == "preview":
            try:
                url = classroom.write_preview(item, os.path.join(core.BASE, "_preview"))
                # 旧アプリは /draft/… で見せていた。ここではプレビューの配り口から見せる
                preview_url = reverse("manage:hp_preview_file", kwargs={"path": url[len("/draft/"):].lstrip("/")})
            except Exception as ex:
                messages.error(request, "プレビューを作れませんでした: %s" % ex)
        elif 何 == "addimg":
            f = request.FILES.get("img_file")
            if not f:
                messages.error(request, "失敗: ファイルが受け取れませんでした。")
            else:
                try:
                    item["images"].append(site_images.写真を足す(f.read(), f.name, "classroom", 1200))
                    messages.success(request, "写真を追加しました。")
                except Exception as ex:
                    messages.error(request, "失敗: %s" % ex)
        elif 何.startswith("repimg:"):
            ri = _番号(何)
            f = request.FILES.get("rep_file_%d" % ri)
            if not (0 <= ri < len(item["reports"])) or not f:
                messages.error(request, "失敗: ファイルが受け取れませんでした。")
            else:
                try:
                    item["reports"][ri]["images"].append(site_images.写真を足す(f.read(), f.name, "classroom", 1200))
                    messages.success(request, "写真を追加しました。")
                except Exception as ex:
                    messages.error(request, "失敗: %s" % ex)
        elif 何.startswith("imgins:"):
            k = _番号(何)
            if 0 <= k < len(item["images"]):
                item["blocks"].append({"t": "img", "v": item["images"][k]})
                messages.success(request, "概要の最後に差し込みました。")
        elif 何.startswith("imgdel:"):
            k = _番号(何)
            if 0 <= k < len(item["images"]):
                f = item["images"].pop(k)
                item["blocks"] = [b for b in item["blocks"] if not (b["t"] == "img" and b["v"] == f)]
        elif 何.startswith("blockadd:"):
            t = 何.split(":", 1)[1]
            if t in ("p", "h3", "list", "note"):
                item["blocks"].append({"t": t, "v": "" if t in ("p", "h3") else [""]})
        elif 何.startswith("blockdel:"):
            bi = _番号(何)
            if 0 <= bi < len(item["blocks"]):
                item["blocks"].pop(bi)
        elif 何.startswith("repadd:"):
            kind = 何.split(":", 1)[1] or "過去の開催"
            item["reports"].insert(0, {"title": "", "date": "", "date_disp": "", "kind": kind, "text": "", "images": []})
            messages.success(request, "開催予定を作りました。日付とタイトルを入れてください。" if kind == "開催予定"
                             else "開催した回を作りました。日付とタイトルを入れてください。")
        elif 何.startswith("repdel:"):
            ri = _番号(何)
            if 0 <= ri < len(item["reports"]):
                item["reports"].pop(ri)
        elif 何.startswith("repmove:"):
            ri = _番号(何)
            j = ri + (1 if 何.endswith(":down") else -1)
            a = item["reports"]
            if 0 <= ri < len(a) and 0 <= j < len(a):
                a[ri], a[j] = a[j], a[ri]
        elif 何.startswith("repimgdel:"):
            ri, k = _番号(何), _番号(何, 2)
            if 0 <= ri < len(item["reports"]) and 0 <= k < len(item["reports"][ri]["images"]):
                item["reports"][ri]["images"].pop(k)
        elif 何 == "repsort":
            # 新しい順（日付の新しいものが上）。日付を入れていないものは、いまの並びのまま下にまとめる
            # （reverse でも同じ日付どうしの並びは変わらない。空の日付はいちばん小さいので下に来る）
            item["reports"].sort(key=lambda r: (r.get("date") or "").strip(), reverse=True)
            messages.success(request, "新しい順に並べ替えました。「保存してページを作り直す」で確定します。")
    else:
        item = cs[i]

    label = "新しいお教室" if (request.GET.get("new") and not _名前(item)) else \
        "お教室の編集：" + (item.get("category") or item.get("title") or "")[:30]
    ctx = {"log": views._記録を取る(request), "i": i, "item": item, "label": label,
           "cats": cats, "kinds": kinds, "preview_url": preview_url, "dirty": dirty}
    ctx.update(_画面用(item))
    return render(request, "manage/hp_class_edit.html", ctx)


def _番号(何, n=1):
    try:
        return int(何.split(":")[n])
    except (IndexError, ValueError):
        return -1

