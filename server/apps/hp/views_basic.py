"""公式サイトの「🏥 基本情報」。中身は mayumi-site/admin/app.py の sec_basic_page と同じ。

旧アプリは画面の入力を全部集めて content.json 全体を保存していた（collect → /api/save）。
ここでは**カードごとに自分の部分だけ**を core.load() → 変える → core.save() で保存する。
ほかの画面（お知らせ・写真など）が同時に直していても、その部分を上書きしない。
"""

import os
import re

from django.contrib import messages
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import repo, site_images, views

# 旧アプリの VALUE_LABELS（並びも同じ）
VALUE_LABELS = [
    ("tel", "電話番号"),
    ("line_url", "LINE予約のURL"),
    ("instagram_url", "InstagramのURL"),
    ("facebook_url", "FacebookのURL"),
    ("blog_url", "ブログのURL"),
    ("calendar_url", "診療カレンダーのURL"),
    ("address_long", "住所（正式表記）"),
    ("address_short", "住所（短縮表記）"),
    ("hours_cta", "診療時間（1行表記・各ページ下部に出るもの）"),
    ("tagline", "ページ上部のオレンジ帯の文"),
    ("footer_lead", "フッターの案内文"),
]

# 旧アプリの PRICE_LABELS
PRICE_LABELS = {
    "bonyu-massage": "母乳外来：乳房マッサージ",
    "bonyu-oushin": "母乳外来：往診",
    "bonyu-soudan": "母乳外来：ご相談",
    "bonyu-all": "母乳外来：まとめた表（料金ページ用）",
    "femcare": "産後ケア：フェムケア",
    "itothermy": "イトオテルミー温熱療法",
    "acupuncture-day": "鍼灸：診療日",
    "acupuncture": "鍼灸：コース",
}

# 旧アプリの BASIC_PARTS（目次の並び。最後は「見出しを足すか」）
BASIC_PARTS = [
    ("hours", "🕒", "診療時間", False),
    ("prices", "💰", "料金表", True),
    ("sns", "🔖", "SNS・ブログ", False),
    ("doc", "📄", "産後ケアの資料", False),
    # 見出し帯の写真。旧アプリでは「写真」タブの先頭にあったが、content.json の
    # page_heads はこの画面で受け持つことになったので、いちばん下に置く。
    ("page_heads", "🏞️", "見出し帯の写真", False),
]

# 旧アプリの snsAdd が足す新しいリンク
SNS_NEW = {"name": "新しいリンク", "url": "https://", "icon": "", "icon_w": 44,
           "style": "icon", "btn_class": "sns__other", "btn_label": "新しいリンク", "show": True}

保存した = "下書きを保存しました（サイトにはまだ反映していません）。"


def _int(値, 既定):
    """旧アプリの parseInt(x,10)||既定 と同じ。"""
    try:
        return int(str(値).strip())
    except (TypeError, ValueError):
        return 既定


def _img_dir(core):
    return os.path.join(core.SITE, "assets", "img")


def _写真の一覧(core):
    """見出し帯に選べる写真（assets/img 直下。SNS のロゴは除く）。"""
    d = _img_dir(core)
    if not os.path.isdir(d):
        return []
    return sorted(fn for fn in os.listdir(d)
                  if fn.lower().endswith((".webp", ".jpg", ".jpeg", ".png")) and not fn.startswith("sns-"))


def _料金表(core, d):
    """料金表の一覧。まとめた表（merge）は見出しだけ直せる。"""
    out = []
    for key, p in d["prices"].items():
        item = {"key": key, "title": PRICE_LABELS.get(key, key), "caption": p.get("caption", ""),
                "note": p.get("note", ""), "merge": bool(p.get("merge"))}
        if p.get("merge"):
            item["names"] = "・".join(PRICE_LABELS.get(k, k) for k in p["merge"])
            item["rows"] = core.merged_price(d, key)["rows"]
        else:
            head = p.get("head") or ["項目", "料金"]
            rows = p.get("rows") or []
            ncol = len(rows[0]) if rows else len(head)
            item["head"] = head[:ncol]
            item["ncol"] = ncol
            item["rows"] = [[(j, c) for j, c in enumerate(row)] for row in rows]
        out.append(item)
    return out


def _資料(core, d):
    """産後ケアの資料（PDF）のいまの状態。"""
    c = ((d.get("docs") or {}).get("bijiris")) or {}
    fn = (c.get("file") or "").strip()
    path = os.path.join(core.SITE, "assets", "doc", fn) if fn else ""
    if path and os.path.isfile(path):
        state = "ok"
        size = os.path.getsize(path) / 1048576
    elif fn:
        state, size = "missing", 0
    else:
        state, size = "none", 0
    return {"show": c.get("show", True), "label": c.get("label", ""), "note": c.get("note", ""),
            "file": fn, "state": state, "size": "%.1f" % size}


@owner_required
def hp_basic(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    blog = repo.部品("blog")
    d = core.load()
    img_dir = _img_dir(core)
    sns = []
    for i, s in enumerate(d.get("sns", [])):
        icon = s.get("icon", "")
        sns.append({"i": i, "name": s.get("name", ""), "url": s.get("url", ""), "icon": icon,
                    "icon_exists": bool(icon) and os.path.exists(os.path.join(img_dir, icon)),
                    "icon_w": s.get("icon_w", 44), "btn_label": s.get("btn_label", ""),
                    "show": s.get("show", True)})
    ph = d.get("page_heads") or {}
    imgs = ph.get("images") or {}
    pics = _写真の一覧(core)
    heads = [{"key": k, "label": label, "cur": imgs.get(k, "")} for k, label in core.PAGE_HEAD_KEYS]
    embed = d.get("sns_embed") or {}
    bt = d.get("blog_top") or {}
    return render(request, "manage/hp_basic.html", {
        "parts": BASIC_PARTS,
        "values": [(k, label, d["values"].get(k, "")) for k, label in VALUE_LABELS],
        "hours_rows": list(enumerate(d["hours"]["rows"])),
        "hours_note": d["hours"].get("note", ""),
        "hours_cta": d["values"].get("hours_cta", ""),
        "prices": _料金表(core, d),
        "sns": sns,
        "sns_embed": {"show": embed.get("show", True), "url": embed.get("url", ""),
                      "height": embed.get("height", 560), "title": embed.get("title", "")},
        "blog_top": {"show": bt.get("show", True), "count": bt.get("count", 6), "height": bt.get("height", 560),
                     "n": len(blog.load().get("posts", []))},
        "doc": _資料(core, d),
        "page_heads": {"show": ph.get("show", True), "veil": ph.get("veil", 65), "rows": heads, "pics": pics},
    })


def _戻る(part=""):
    return redirect(reverse("manage:hp_basic") + ("#sec-" + part if part else ""))


@owner_required
@require_POST
def hp_basic_save(request):
    """カードごとの保存。part で「どの部分か」を受け、その部分だけ content.json に書く。"""
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    P = request.POST
    part = P.get("part", "")
    d = core.load()

    if part == "basic":
        for k, _label in VALUE_LABELS:
            if "v-" + k in P:
                d["values"][k] = P.get("v-" + k, "")
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("basic-top")

    if part == "hours":
        rows = d["hours"]["rows"]
        for i, row in enumerate(rows):
            for j in range(len(row)):
                if "h-%d-%d" % (i, j) in P:
                    row[j] = P.get("h-%d-%d" % (i, j), "")
        if "hnote" in P:
            d["hours"]["note"] = P.get("hnote", "")
        if "v-hours_cta" in P:
            d["values"]["hours_cta"] = P.get("v-hours_cta", "")
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("hours")

    if part in ("price", "price_row_add", "price_row_del"):
        key = P.get("key", "")
        if key not in d["prices"]:
            messages.error(request, "その料金表はありません")
            return _戻る("prices")
        p = d["prices"][key]
        p["caption"] = (P.get("caption") or "").strip()
        p["note"] = (P.get("note") or "").strip()
        if not p.get("merge"):
            # 画面の入力をそのまま行にする（旧アプリの collect と同じ。空の行も残す）
            nrows, ncol = _int(P.get("nrows"), 0), _int(P.get("ncol"), 2)
            p["rows"] = [[P.get("cell-%d-%d" % (i, j), "") for j in range(ncol)] for i in range(nrows)]
            if part == "price_row_add":
                p["rows"].append([""] * ncol)
            elif part == "price_row_del":
                i = _int(P.get("row"), -1)
                if 0 <= i < len(p["rows"]):
                    p["rows"].pop(i)
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("price-" + key)

    if part == "sns_add":
        d.setdefault("sns", []).append(dict(SNS_NEW))
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("sns")

    if part in ("sns", "sns_del", "sns_move"):
        i = _int(P.get("i"), -1)
        sns = d.setdefault("sns", [])
        if not (0 <= i < len(sns)):
            messages.error(request, "そのリンクは見つかりませんでした。")
            return _戻る("sns")
        s = sns[i]
        # 旧アプリの snsDel / snsMove も、先に画面の入力を集めてから動かす
        for k in ("name", "url", "btn_label"):
            if "s-" + k in P:
                s[k] = P.get("s-" + k, "")
        if "s-icon_w" in P:
            s["icon_w"] = _int(P.get("s-icon_w"), 44)
        if "s-show" in P:
            s["show"] = P.get("s-show") == "1"
        if part == "sns_del":
            sns.pop(i)
        elif part == "sns_move":
            j = i + (_int(P.get("dir"), 0))
            if 0 <= j < len(sns):
                sns[i], sns[j] = sns[j], sns[i]
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("sns")

    if part == "sns_icon":
        # ロゴ画像の差し替え（旧アプリの /api/replace-image と同じ）
        target = os.path.basename((P.get("target") or "").strip())
        f = request.FILES.get("file")
        if target and f:
            try:
                site_images.写真を差し替える(target, f.read())
                messages.success(request, "%s を差し替えました。" % target)
            except Exception as ex:
                messages.error(request, "画像の保存に失敗しました: %s" % ex)
        elif not target:
            messages.error(request, "このリンクにはロゴのファイル名がありません。先に名前を決めてから入れてください。")
        return _戻る("sns")

    if part == "sns_embed":
        c = d.setdefault("sns_embed", {})
        c["show"] = P.get("fb-show") == "1"
        c["url"] = P.get("fb-url", "")
        c["height"] = _int(P.get("fb-height"), 560)
        c["title"] = P.get("fb-title", "")
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("sns-embed")

    if part == "blog_top":
        c = d.setdefault("blog_top", {})
        c["show"] = P.get("bt-show") == "1"
        c["count"] = _int(P.get("bt-count"), 6)
        c["height"] = _int(P.get("bt-height"), 560)
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("blog-top")

    if part == "doc":
        c = d.setdefault("docs", {}).setdefault("bijiris", {})
        c["show"] = P.get("doc-show") == "1"
        c["label"] = P.get("doc-label", "")
        c["note"] = P.get("doc-note", "")
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("doc")

    if part == "doc_upload":
        # 旧アプリの handle_doc_upload と同じ。日本語のファイル名は URL で扱いにくいので決まった名前にする
        key = re.sub(r"[^a-z0-9_-]", "", (P.get("key") or "bijiris").lower()) or "bijiris"
        f = request.FILES.get("file")
        if f:
            name = "%s-text.pdf" % key
            dst_dir = os.path.join(core.SITE, "assets", "doc")
            os.makedirs(dst_dir, exist_ok=True)
            with open(os.path.join(dst_dir, name), "wb") as out:
                for chunk in f.chunks():
                    out.write(chunk)
            d.setdefault("docs", {}).setdefault(key, {})["file"] = name
            core.save(d)
            messages.success(request, "PDF を入れ替えました（%s）。" % name)
        else:
            messages.error(request, "ファイルが受け取れませんでした。")
        return _戻る("doc")

    if part == "page_heads":
        c = d.setdefault("page_heads", {})
        imgs = c.setdefault("images", {})
        for k, _label in core.PAGE_HEAD_KEYS:
            if "ph-" + k in P:
                imgs[k] = P.get("ph-" + k, "")
        c["show"] = P.get("ph-show") == "1"
        c["veil"] = max(0, min(100, _int(P.get("ph-veil"), 65)))
        core.save(d)
        messages.success(request, 保存した)
        return _戻る("page_heads")

    messages.error(request, "何を保存するのか分かりませんでした。")
    return _戻る()


@owner_required
@require_POST
def hp_basic_price_preview(request):
    """入力した内容で料金表の見え方だけを作る（旧アプリの /api/price-preview）。content.json は書き換えない。"""
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    P = request.POST
    d = core.load()
    key = P.get("key", "")
    if key not in d["prices"]:
        return HttpResponse("その料金表はありません", status=404)
    p = dict(d["prices"][key])
    p["caption"] = (P.get("caption") or "").strip()
    p["note"] = (P.get("note") or "").strip()
    if p.get("merge"):
        d["prices"] = dict(d["prices"])
        d["prices"][key] = p
        p = core.merged_price(d, key)
    else:
        nrows, ncol = _int(P.get("nrows"), 0), _int(P.get("ncol"), 2)
        p["rows"] = [[P.get("cell-%d-%d" % (i, j), "") for j in range(ncol)] for i in range(nrows)]
    return render(request, "manage/hp_price_preview.html", {
        "title": PRICE_LABELS.get(key, key), "table": core.render_price(p),
    })
