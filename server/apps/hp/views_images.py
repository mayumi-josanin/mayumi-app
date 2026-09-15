"""公式サイトの「🖼️ 写真」。中身は mayumi-site/admin/app.py の sec_images（＋ sec_page_heads）と同じ。

- 写真の一覧と差し替え（/api/replace-image）
- どのページからも使われていない写真を消す（/api/del-unused）
- 見出し帯の写真（page_heads。旧アプリでは「写真」タブの中にあり、「下書き保存」で content.json に入る）

写真の一覧を作る2つ（image_files / used_images）は app.py にしか無いので、ここに小さく写した。
"""

import os
import re
import shutil

from django.contrib import messages
from django.shortcuts import redirect, render
from django.urls import reverse

from apps.manage.permissions import owner_required

from . import repo, site_images, views

# 見出し帯に選べる写真の種類（sec_page_heads と同じ）
_見出し帯に使える = (".webp", ".jpg", ".jpeg", ".png")


def _img_dir():
    core = repo.部品("core")
    return os.path.join(core.SITE, "assets", "img")


def 写真の一覧():
    """image_files: assets/img の直下にあるファイルだけ。

    blog / news / event / classroom といった入れ物のフォルダは**入れない**。
    旧アプリで以前フォルダ名まで拾っていて、「未使用の写真を消す」がフォルダごと片づけ、
    つぶやきの写真1002枚がまとめて消えたことがある。
    """
    img_dir = _img_dir()
    if not os.path.isdir(img_dir):
        return []
    return sorted(fn for fn in os.listdir(img_dir)
                  if not fn.startswith(".") and os.path.isfile(os.path.join(img_dir, fn)))


def 使われている写真():
    """used_images: サイトのHTMLとCSSから、実際に使われている写真の名前を集める。"""
    core = repo.部品("core")
    used = set()
    targets = list(core.site_files())
    css = os.path.join(core.SITE, "css", "style.css")
    if os.path.exists(css):
        targets.append(css)
    for p in targets:
        s = open(p, encoding="utf-8").read()
        for m in re.finditer(r"assets/img/(?:[\w-]+/)?([A-Za-z0-9_.\-]+)", s):
            used.add(m.group(1))
    return used


def 未使用の写真を消す():
    """/api/del-unused: 使われていない写真を admin/backups/_images へ移す（消さずに控える）。"""
    core = repo.部品("core")
    img_dir = _img_dir()
    used = 使われている写真()
    gone = []
    bak = os.path.join(core.BACKUP_DIR, "_images")
    os.makedirs(bak, exist_ok=True)
    for fn in 写真の一覧():
        if fn not in used:
            shutil.move(os.path.join(img_dir, fn), os.path.join(bak, fn))
            gone.append(fn)
    return [
        "%d枚を削除しました：%s" % (len(gone), "、".join(gone)) if gone else "未使用の写真はありませんでした。",
        "（念のため backups/_images に移してあります）" if gone else "",
    ]


def _見出し帯の候補(files):
    """sec_page_heads: 見出し帯に選べる写真。SNS のロゴ（sns-）は除く。"""
    return [fn for fn in files if fn.lower().endswith(_見出し帯に使える) and not fn.startswith("sns-")]


def _見出し帯を保存(request, files):
    """旧アプリの collect() の page_heads の部分。画面の値を content.json に入れる。"""
    core = repo.部品("core")
    d = core.load()
    c = d.get("page_heads") or {}
    imgs = c.get("images") or {}
    候補 = set(_見出し帯の候補(files))
    for key, _label in core.PAGE_HEAD_KEYS:
        v = (request.POST.get("ph_" + key) or "").strip()
        # 選べるのは一覧にある写真だけ（無い名前を入れると、帯が壊れた写真になる）
        imgs[key] = v if v in 候補 else ""
    c["images"] = imgs
    c["show"] = request.POST.get("ph_show", "1") == "1"
    try:
        veil = int(request.POST.get("ph_veil", "") or 65)
    except ValueError:
        veil = 65
    c["veil"] = max(0, min(100, veil))
    d["page_heads"] = c
    core.save(d)


@owner_required
def hp_images(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    core = repo.部品("core")
    if request.method == "POST":
        何 = request.POST.get("action", "")
        if 何 == "replace":
            target = os.path.basename((request.POST.get("target") or "").strip())
            f = request.FILES.get("file")
            if target not in 写真の一覧():
                messages.error(request, "差し替える写真が見つかりませんでした: %s" % target)
            elif f is None:
                messages.error(request, "画像が選ばれていません。")
            else:
                try:
                    site_images.写真を差し替える(target, f.read())
                    messages.success(request, "写真を差し替えました: %s" % target)
                except Exception as ex:
                    messages.error(request, "画像の保存に失敗しました: %s" % ex)
        elif 何 == "del_unused":
            views._記録を出す(request, [l for l in 未使用の写真を消す() if l], "未使用の写真を消す")
        elif 何 == "save_heads":
            _見出し帯を保存(request, 写真の一覧())
            messages.success(request, "下書きを保存しました（サイトにはまだ反映していません）。")
        return redirect("manage:hp_images")

    files = 写真の一覧()
    used = 使われている写真()
    img_dir = _img_dir()
    cards = []
    for fn in files:
        cards.append({
            "name": fn,
            "kb": os.path.getsize(os.path.join(img_dir, fn)) // 1024,
            "used": fn in used,
            "src": reverse("manage:hp_site_file", kwargs={"path": "assets/img/" + fn}),
        })
    unused = [fn for fn in files if fn not in used]

    d = core.load()
    c = d.get("page_heads") or {}
    imgs = c.get("images") or {}
    pics = _見出し帯の候補(files)
    rows = []
    for key, label in core.PAGE_HEAD_KEYS:
        cur = imgs.get(key, "")
        rows.append({
            "key": key, "label": label, "cur": cur,
            "src": reverse("manage:hp_site_file", kwargs={"path": "assets/img/" + cur}) if cur else "",
        })
    return render(request, "manage/hp_images.html", {
        "log": views._記録を取る(request),
        "cards": cards,
        "unused": unused,
        "unused_names": "、".join(unused[:8]) + ("…" if len(unused) > 8 else ""),
        "heads": {"show": c.get("show", True), "veil": c.get("veil", 65), "rows": rows, "pics": pics},
    })
