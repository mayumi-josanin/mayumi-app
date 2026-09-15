"""公式サイトの「🎛️ 選択肢の設定」。中身は mayumi-site/admin/app.py の sec_options と同じ。

旧アプリでは、画面の入力を全部集めて（collect）から 追加／削除／並べ替え をして
content.json 全体を保存していた。ここでも同じ順で、
  1. 画面の入力（名前・色・幅）を content.json の options に写す
  2. 押されたボタンの操作（追加・削除・上へ／下へ）をする
  3. 保存する
とするので、書きかけの名前が操作のたびに消えない。
"""

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils.safestring import mark_safe

from apps.manage.permissions import owner_required

from . import repo, views

# 4種類の選択肢（content.json の options のキー）と、画面での出し方
種類 = ("news_categories", "class_categories", "class_kinds", "widths")

# 「＋ 追加」で増えるときの中身（app.py の OPT_NEW と同じ。
#  お教室の分類は旧アプリの OPT_NEW に無く、名前も色も空のまま増えていたので、ここでは名前を添える）
新しい項目 = {
    "news_categories": {"name": "新しい種別", "color": "#8a9e7e"},
    "class_categories": {"name": "新しいお教室", "color": "#8a9e7e"},
    "class_kinds": {"name": "新しい分類", "color": "#8a9e7e"},
    "widths": {"label": "新しい幅", "value": "50%"},
}

# 名前が無いときのラベルの色（core.cat_html と同じ）
既定の色 = "#9a8070"

# 保存したときの文言（旧アプリの /api/save の返事と同じ）
保存した = "下書きを保存しました（サイトにはまだ反映していません）。"

# 画面の見出し・説明（旧アプリの文言と同じ。「手元に反映」だけは、この画面に無いので「保存」に言い換えた）
画面の説明 = [
    {
        "key": "news_categories", "title": "📢 お知らせの種別", "add": "＋ 種別を追加",
        "hint": mark_safe(
            "お知らせのタイトルの前に出るラベルです。色は四角を押すと選べます。<br>"
            "<b>名前を変えると、すでにその種別を付けたお知らせは「なし」扱いになります。</b>"
            "名前を変えたときは、お知らせ側も選び直してください。"),
    },
    {
        "key": "class_categories", "title": "🎒 お教室の正式な名称", "add": "＋ お教室の名称を追加",
        "hint": mark_safe(
            "お教室のタイトルの前に出るラベルです（「梅干しづくり教室」など）。<br>"
            "サイトに載っている「梅干し作りの会」などは<b>その回のタイトル</b>で、"
            "正式なお教室の名前はここで決めます。同じ教室を何回開いても、"
            "同じ名称を選べば同じラベルで並びます。<br>"
            "<b>名前を変えると、すでにその名称を付けたお教室は「なし」扱いになります。</b>"),
    },
    {
        "key": "class_kinds", "title": "🗂 お教室の分類（定期開催・臨時開催）", "add": "＋ 分類を追加",
        "hint": mark_safe(
            "お教室のページで、見に来た方が<b>この分類でしぼれます</b>。"
            "色は、しぼり込みボタンの丸い印に使われます。<br>"
            "<b>名前を変えると、すでにその分類を付けたお教室は「なし」扱いになります。</b>"
            "名前を変えたときは、お教室側も選び直してください。"),
    },
]


def _選択肢を取り出す(core, d):
    """content.json の options を、既定値も埋めた形で返す（core.options と同じ中身）。

    旧アプリは、入っていない種類は既定値を画面に出しつつ、その名前の変更は捨てていた
    （collect が d.options に無い種類を飛ばすため）。ここでは既定値を options に写してから
    直すので、既定値のままの種類も名前や色を変えられる。
    """
    o = core.options(d)
    d["options"] = {k: [dict(x) for x in (o.get(k) or [])] for k in 種類}
    return d["options"]


def _入力を写す(post, o):
    """画面の入力を options に写す（旧アプリの collect() の options の部分と同じ）。"""
    for key in 種類:
        欄 = ("label", "value") if key == "widths" else ("name", "color")
        for i, item in enumerate(o[key]):
            for f in 欄:
                v = post.get("oc__%s__%d__%s" % (key, i, f))
                if v is not None:
                    item[f] = v


def _操作する(action, o):
    """押されたボタンの操作。add|種類 / del|種類|番号 / move|種類|番号|向き。"""
    部 = (action or "").split("|")
    何 = 部[0] if 部 else ""
    key = 部[1] if len(部) > 1 and 部[1] in 種類 else None
    if not key:
        return
    lst = o[key]
    if 何 == "add":
        lst.append(dict(新しい項目[key]))
    elif 何 == "del" and len(部) > 2:
        i = _番号(部[2])
        if i is not None and 0 <= i < len(lst):
            del lst[i]
    elif 何 == "move" and len(部) > 3:
        i, 向き = _番号(部[2]), _番号(部[3])
        if i is None or 向き is None:
            return
        j = i + 向き
        if 0 <= i < len(lst) and 0 <= j < len(lst):
            lst[i], lst[j] = lst[j], lst[i]


def _番号(s):
    try:
        return int(s)
    except (TypeError, ValueError):
        return None


@owner_required
def hp_options(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    if request.method == "POST":
        d = core.load()
        o = _選択肢を取り出す(core, d)
        _入力を写す(request.POST, o)
        _操作する(request.POST.get("action", "save"), o)
        core.save(d)
        messages.success(request, 保存した)
        return redirect("manage:hp_options")

    o = _選択肢を取り出す(core, core.load())
    groups = []
    for g in 画面の説明:
        items = [{"name": c.get("name", ""), "color": c.get("color") or 既定の色} for c in o[g["key"]]]
        groups.append(dict(g, items=items))
    return render(request, "manage/hp_options.html", {
        "groups": groups,
        "widths": [{"label": w.get("label", ""), "value": w.get("value", "")} for w in o["widths"]],
    })
