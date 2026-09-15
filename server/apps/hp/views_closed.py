"""公式サイトの「📅 診療カレンダー」。中身は mayumi-site/admin/app.py の sec_closed と同じ。

旧アプリは JS で /api/sched-new・sched-save・sched-del・sched-image・calendar・
app-calendar・app-calendar-import を叩いていた。ここでは同じことを
1つの画面の POST（action=new / save / del / import）と GET の絞り込みで行う。

アプリのカレンダーからの取り込みは、旧アプリでは GAS（getCalendar）を読んでいたが、
ここでは**同じサーバーの表（CalendarEvent）から、GAS が返していたのと同じ形**で組む。
"""

import calendar as _cal
import datetime
import re

from django.contrib import messages
from django.shortcuts import redirect, render
from django.urls import reverse

from apps.manage.permissions import owner_required

from . import repo, site_images, views

# ---- app.py から写したもの（アプリのカレンダーの読み方）----------------------
# アプリのカレンダーから持ってくる種類。**イベントは持ってこない。**
# お教室の開催日が自動で入るので、二重に並ぶのを避けるため。
APP_CAL_KINDS = {"closed": "休診", "visit": "往診", "postpartum": "訪問産後ケア"}

# 種類の見分け方。カテゴリ → 題名 → 色 の順に見る。
# 「訪問産後ケア」なのに休診の色が付いた行が実際にあるので、題名を色より先に見る。
APP_CAL_BY_CATEGORY = {"休診": "closed", "往診": "visit",
                       "訪問産後ケア": "postpartum", "訪問型産後ケア": "postpartum"}
APP_CAL_BY_COLOR = {"#e57373": "closed", "#ffc7fa": "visit"}


def app_cal_kind(ev):
    """アプリの1行が、どの種類にあたるか。当てはまらなければ None。（app.py と同じ）"""
    cat = str(ev.get("category") or "").strip()
    if cat in APP_CAL_BY_CATEGORY:
        return APP_CAL_BY_CATEGORY[cat]
    title = str(ev.get("title") or "").strip()
    if title in APP_CAL_BY_CATEGORY:
        return APP_CAL_BY_CATEGORY[title]
    if "産後ケア" in title:
        return "postpartum"
    if "往診" in title:
        return "visit"
    if "休診" in title:
        return "closed"
    return APP_CAL_BY_COLOR.get(str(ev.get("color") or "").strip().lower())


def app_news_date(raw):
    """アプリが返す日付を "2026-09-07" の形にする。（app.py と同じ）"""
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


def _アプリのカレンダー():
    """旧アプリが GAS の getCalendar から受け取っていた payload と同じものを、サーバーの表から組む。

    お客様アプリに出しているのと同じ条件（消していない・公開中・予約公開は時刻が来てから）
    なので、gasapi の getCalendar をそのまま使う。項目（date / title / category / color）も同じ。
    """
    from apps.gasapi.views import _カレンダー

    return _カレンダー()


def fetch_app_calendar(d):
    """アプリのカレンダーから、休診・往診・訪問産後ケアを取ってくる。（app.py と同じ。読む先だけ表）

    取り込み済みかどうかは、同じ種類の予定にその日付が入っているかで見る。
    """
    calendar_block = repo.部品("calendar_block")
    payload = _アプリのカレンダー()
    have = {}                       # 種類 → すでに入っている日付
    for c in d.get("schedule", []):
        k = (c.get("kind") or "closed").strip()
        have.setdefault(k, set()).update(calendar_block.dates_of(c))

    out = []
    seen = set()
    for ev in payload.get("events", []):
        kind = app_cal_kind(ev)
        if kind not in APP_CAL_KINDS:
            continue
        date = app_news_date(ev.get("date"))
        if not date or (kind, date) in seen:
            continue                # 同じ種類・同じ日が二重に入っていることがある
        seen.add((kind, date))
        out.append({
            "key": "%s|%s" % (kind, date),
            "date": date,
            "kind": kind,
            "label": APP_CAL_KINDS[kind],
            "title": " ".join(str(ev.get("title") or "").split()),
            "done": date in have.get(kind, set()),
        })
    out.sort(key=lambda x: (x["date"], x["kind"]))
    return out


def import_app_calendar(keys):
    """選ばれた日を、診療カレンダーの予定に足す。（app.py と同じ）

    種類ごとに1つの予定へまとめる。サイト側は1つの予定に何日でも
    登録できる作りなので、日ごとに予定を作るとかえって扱いにくい。
    """
    core = repo.部品("core")
    calendar_block = repo.部品("calendar_block")
    d = core.load()
    want = set(keys or [])
    items = [x for x in fetch_app_calendar(d) if x["key"] in want and not x["done"]]
    if not items:
        return ["取り込むものがありませんでした。"], 0

    sched = d.setdefault("schedule", [])
    log = []
    added = 0

    by_kind = {}
    for x in items:
        by_kind.setdefault(x["kind"], []).append(x["date"])

    for kind, dates in by_kind.items():
        label = APP_CAL_KINDS[kind]
        # 同じ種類・同じ名前の予定があれば、そこに日付を足す
        target = None
        for c in sched:
            if (c.get("kind") or "closed") == kind and (c.get("title") or "").strip() == label:
                target = c
                break
        if target is None:
            target = {"id": _次の番号(sched), "kind": kind, "dates": [], "title": label,
                      "note": "", "text": "", "images": [], "show": True}
            sched.append(target)
            log.append("「%s」の予定を作りました。" % label)
        got = set(calendar_block.dates_of(target)) | set(dates)
        target.pop("date", None)
        target["dates"] = sorted(got)
        added += len(dates)
        log.append("%s に %d日を足しました（ぜんぶで %d日）。"
                   % (label, len(dates), len(got)))

    core.save(d)
    log.insert(0, "%d日を取り込みました。" % added)
    return log, added


# ---- 予定の作り方（app.py の /api/sched-* と同じ）-----------------------------
def _次の番号(sched, 除く=None):
    """予定ごとのページのURLに使う番号 sc01, sc02…。一度決めたら変えない。"""
    used = {(x.get("id") or "").strip() for j, x in enumerate(sched) if j != 除く}
    k = 1
    while ("sc%02d" % k) in used:
        k += 1
    return "sc%02d" % k


def _新しい予定(sched):
    return {"id": _次の番号(sched), "kind": "closed", "dates": [], "title": "",
            "note": "", "text": "", "images": [], "show": True}


def _カレンダーを作り直す(d, log):
    """保存のあと、サイトのカレンダーと予定ごとのページを作り直す（/api/sched-save と同じ）。"""
    core = repo.部品("core")
    calendar_block = repo.部品("calendar_block")
    try:
        core.apply_blocks(d)                     # カレンダーを作り直す
        n = calendar_block.build_pages()
        calendar_block.update_sitemap()
        log.append("カレンダーに %d件の予定を入れました（詳しいページ %d件）。"
                   % (len(calendar_block.collect(d)), n))
    except Exception as ex:
        log.append("カレンダーの作り直しに失敗: %s" % ex)


def _画面の予定(request):
    """画面の入力から予定を組む（JS の scGrab と同じ）。"""
    P = request.POST
    return {
        "kind": P.get("kind", "closed"),
        "title": P.get("title", ""),
        "note": P.get("note", "").strip(),
        "text": P.get("text", ""),
        "show": P.get("show", "1") == "1",
        "dates": sorted({x for x in P.getlist("dates") if x}),
        "images": [x for x in P.getlist("images") if x],
    }


# ---- 画面 -------------------------------------------------------------------
def _整数(s, 既定=-1):
    try:
        return int(s)
    except (TypeError, ValueError):
        return 既定


def _写真の場所(f):
    return reverse("manage:hp_site_file", kwargs={"path": "assets/img/calendar/" + f})


def _月の升目(y, m, evs, kinds, sel):
    """「🗓️ カレンダーで見る」の升目（旧アプリの calvRender と同じ並び）。"""
    by_day = {}
    for e in evs:
        by_day.setdefault(e["d"], []).append(e)
    first = (_cal.monthrange(y, m)[0] + 1) % 7          # 1日の曜日（0=日）
    last = _cal.monthrange(y, m)[1]
    today = datetime.date.today().isoformat()
    cells = [None] * first
    for day in range(1, last + 1):
        ds = "%04d-%02d-%02d" % (y, m, day)
        d_evs = by_day.get(ds, [])
        dow = (first + day - 1) % 7
        cells.append({
            "n": day, "ds": ds, "dow": dow,
            "today": ds == today, "sel": ds == sel,
            # マスに入りきらない分は「＋n」でまとめる。押せば下に全部出る
            "marks": [dict(kinds.get(e["k"], {"m": "・", "c": "#9a8070", "w": 0}), t=e["t"]) for e in d_evs[:4]],
            "more": max(0, len(d_evs) - 4),
        })
    cells += [None] * ((7 - ((first + last) % 7)) % 7)
    return cells


@owner_required
def hp_closed(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    calendar_block = repo.部品("calendar_block")

    if request.method == "POST":
        何 = request.POST.get("action", "")
        d = core.load()
        sc = d.setdefault("schedule", [])
        i = _整数(request.POST.get("index"))

        if 何 == "new":
            sc.insert(0, _新しい予定(sc))
            core.save(d)
            messages.info(request, "日付を選んで「追加」、名前を入れて保存してください。")
            return redirect(reverse("manage:hp_closed") + "?edit=0&new=1")

        if 何 == "save":
            if not (0 <= i < len(sc)):
                messages.error(request, "その位置の予定がありません。")
                return redirect("manage:hp_closed")
            item = dict(sc[i])
            item.pop("date", None)                       # 古い形は残さない
            item.pop("url", None)                        # 行き先は自動で決まる
            item.update(_画面の予定(request))
            log = []
            up = request.FILES.get("photo")
            if up is not None:
                try:
                    name = site_images.写真を足す(up.read(), up.name, "calendar", 1200)
                    item["images"] = item["images"] + [name]
                    log.append("写真を追加しました（%s）。" % name)
                except Exception as ex:
                    log.append("写真の取り込みに失敗: %s" % ex)
            if not item["dates"]:
                # 旧アプリと同じく保存しない。入力は消さずにそのまま画面に戻す
                messages.error(request, "日付が1つも入っていません。")
                return _画面(request, d, edit=i, item=item, log={"title": "予定の編集", "lines": log} if log else None)
            if not (item.get("id") or "").strip():
                item["id"] = _次の番号(sc, 除く=i)
            sc[i] = item
            core.save(d)
            log.insert(0, "保存しました。")
            _カレンダーを作り直す(d, log)
            views._記録を出す(request, log, "保存してカレンダーに反映")
            return redirect(reverse("manage:hp_closed") + "?edit=%d" % i)

        if 何 == "del":
            if not (0 <= i < len(sc)):
                views._記録を出す(request, ["その位置の予定がありません。"], "これを削除")
                return redirect("manage:hp_closed")
            gone = sc.pop(i)
            core.save(d)
            core.apply_blocks(d)
            calendar_block.build_pages()
            calendar_block.update_sitemap()
            views._記録を出す(request, ["「%s」を削除しました。" % (gone.get("title") or "名前なしの予定")], "これを削除")
            return redirect("manage:hp_closed")

        if 何 == "import":
            keys = request.POST.getlist("keys")
            if not keys:
                messages.error(request, "取り込む日が選ばれていません。")
                return redirect(reverse("manage:hp_closed") + "?import=1")
            try:
                log, _added = import_app_calendar(keys)
            except Exception as ex:
                log = ["取り込みに失敗しました: %s" % ex]
            views._記録を出す(request, log, "📥 アプリのカレンダーから取り込む")
            return redirect("manage:hp_closed")

        return redirect("manage:hp_closed")

    return _画面(request, core.load())


def _画面(request, d, edit=None, item=None, log=None):
    calendar_block = repo.部品("calendar_block")
    G = request.GET
    sc = d.get("schedule", [])
    KINDS = calendar_block.KINDS
    labels = {k[0]: k[2] for k in KINDS}
    colors = {k[0]: k[3] for k in KINDS}
    kinds = {k: {"m": mark, "l": label, "c": color, "w": 1 if wide else 0}
             for k, mark, label, color, wide in KINDS}

    # サイトに出るものと同じ中身（calendar_block.collect）。手で入れた予定には
    # 押したときに編集を開けるよう位置（si）を添える。お教室から自動で入るものには付かない。
    evs = calendar_block.collect(d)
    where = {}
    for i, c in enumerate(sc):
        if c.get("show") is False:
            continue
        kind = (c.get("kind") or "closed").strip()
        if kind not in calendar_block.KIND_IDS:
            kind = "event"
        title = (c.get("title") or "").strip() or labels[kind]
        for dt in calendar_block.dates_of(c):
            where.setdefault((dt, kind, title), i)
    for ev in evs:
        ev["si"] = where.get((ev["d"], ev["k"], ev["t"]))
        ev["v"] = kinds.get(ev["k"], {"m": "・", "l": ev["k"], "c": "#9a8070", "w": 0})
        ev["site"] = reverse("manage:hp_site_file", kwargs={"path": ev["u"].lstrip("/")}) if ev.get("u") else ""

    counts = {k: len([e for e in evs if e["k"] == k]) for k in labels}
    legend = [(labels[k], colors[k], counts.get(k, 0)) for k in calendar_block.KIND_IDS]

    # 🗓️ カレンダーで見る
    today = datetime.date.today()
    m = re.match(r"^(\d{4})-(\d{1,2})$", G.get("ym", ""))
    y, mo = (int(m.group(1)), int(m.group(2))) if m and 1 <= int(m.group(2)) <= 12 else (today.year, today.month)
    sel = G.get("day", "")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", sel):
        sel = ""
    prev_m = datetime.date(y, mo, 1) - datetime.timedelta(days=1)
    next_m = datetime.date(y, mo, _cal.monthrange(y, mo)[1]) + datetime.timedelta(days=1)
    day_evs, day_label = [], ""
    if sel:
        day_evs = [e for e in evs if e["d"] == sel]
        yy, mm, dd = (int(x) for x in sel.split("-"))
        try:
            w = "日月火水木金土"[(datetime.date(yy, mm, dd).weekday() + 1) % 7]
            day_label = "%d月%d日（%s）" % (mm, dd, w)
        except ValueError:
            sel, day_evs = "", []

    # 🗓️ 予定の一覧（種類でしぼる）
    f = G.get("kind", "")
    rows = []
    for i, c in enumerate(sc):
        if f and (c.get("kind") or "").strip() != f:
            continue
        ds = calendar_block.dates_of(c)
        kind = c.get("kind") or ""
        rows.append({
            "i": i, "c": c,
            "label": labels.get(kind, kind), "color": colors.get(kind, "#9a8070"),
            "shown": "、".join(ds[:3]) + (" ほか%d日" % (len(ds) - 3) if len(ds) > 3 else ""),
            "ndates": len(ds),
            "detail": bool(c.get("text") or (c.get("images") or [])),
            "hidden": c.get("show") is False,
        })
    count_text = ("%d件 / 全%d件" % (len(rows), len(sc))) if f else ("全%d件" % len(sc))

    # 予定の編集
    if edit is None:
        edit = _整数(G.get("edit"))
    edit_item = None
    if item is not None:
        edit_item = item
    elif 0 <= edit < len(sc):
        edit_item = dict(sc[edit])
        edit_item["dates"] = calendar_block.dates_of(edit_item)
    if edit_item is not None:
        edit_item["images"] = [{"f": x, "src": _写真の場所(x)} for x in (edit_item.get("images") or [])]
        edit_title = "新しい予定" if G.get("new") else "予定の編集：" + (edit_item.get("title") or "")
    else:
        edit, edit_title = -1, ""

    # 📥 アプリのカレンダーから取り込む
    app_items, app_error = None, ""
    if G.get("import"):
        try:
            app_items = fetch_app_calendar(d)
        except Exception as ex:
            app_error = "アプリのカレンダーを読めませんでした: %s" % ex
    app_yet = len([x for x in app_items if not x["done"]]) if app_items else 0

    return render(request, "manage/hp_closed.html", {
        "log": log or views._記録を取る(request),
        "legend": legend, "kind_opts": [(k[0], k[2]) for k in KINDS],
        "site_calendar": reverse("manage:hp_site_file", kwargs={"path": "reception/"}) + "#calendar",
        "ym": "%04d-%02d" % (y, mo), "month_title": "%d年 %d月" % (y, mo),
        "prev_ym": "%04d-%02d" % (prev_m.year, prev_m.month), "next_ym": "%04d-%02d" % (next_m.year, next_m.month),
        "today_ym": "%04d-%02d" % (today.year, today.month),
        "cells": _月の升目(y, mo, evs, kinds, sel), "kinds": kinds.values(),
        "sel": sel, "day_label": day_label, "day_evs": day_evs,
        "rows": rows, "filter": f, "count_text": count_text, "total": len(sc),
        "edit": edit, "edit_item": edit_item, "edit_title": edit_title, "today": today.isoformat(),
        "importing": bool(G.get("import")), "app_items": app_items, "app_error": app_error,
        "app_all": len(app_items or []), "app_yet": app_yet,
    })
