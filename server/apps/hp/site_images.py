"""公式サイトの写真の保存。mayumi-site/admin/app.py の save_sub_image / save_new_image / save_image と同じ。

app.py は HTTP サーバーごと1つのファイルで import できないので、写真の3つだけここに写した。
縮小の幅・WebP の品質・名前の付け方を変えると、旧アプリで入れた写真と見た目が揃わなくなる。
"""

import datetime
import io
import os
import re
import shutil

from . import repo


def _img_dir():
    core = repo.部品("core")
    return os.path.join(core.SITE, "assets", "img")


def 写真を足す(data: bytes, orig_name: str, folder: str, max_w: int = 1200) -> str:
    """save_sub_image: お知らせ・つぶやき・お教室・カレンダー用。assets/img/<folder>/ に WebP で保存。"""
    from PIL import Image

    core = repo.部品("core")
    dst_dir = os.path.join(core.SITE, "assets", "img", folder)
    os.makedirs(dst_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(orig_name or ""))[0]
    stem = re.sub(r"[^A-Za-z0-9_-]", "", stem).strip("-_")
    if not stem:
        stem = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = "%s_%s.webp" % (folder, stem)
    n = 2
    while os.path.exists(os.path.join(dst_dir, name)):
        name = "%s_%s_%d.webp" % (folder, stem, n)
        n += 1
    im = Image.open(io.BytesIO(data))
    if im.mode in ("P", "CMYK", "LA"):
        im = im.convert("RGB")
    if im.mode == "RGBA":
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[3])
        im = bg
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
    im.save(os.path.join(dst_dir, name), "WEBP", quality=84, method=6)
    return name


def 新しい写真(data: bytes, orig_name: str) -> str:
    """save_new_image: サイト全体の写真（assets/img/）を増やす。幅1600・WebP。"""
    from PIL import Image

    img_dir = _img_dir()
    os.makedirs(img_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(orig_name or ""))[0]
    stem = re.sub(r"[^A-Za-z0-9_-]", "", stem).strip("-_")
    if not stem:
        stem = "photo_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = stem + ".webp"
    n = 2
    while os.path.exists(os.path.join(img_dir, name)):
        name = "%s_%d.webp" % (stem, n)
        n += 1
    im = Image.open(io.BytesIO(data))
    if im.mode in ("P", "CMYK"):
        im = im.convert("RGB")
    if im.width > 1600:
        im = im.resize((1600, round(im.height * 1600 / im.width)), Image.LANCZOS)
    im.save(os.path.join(img_dir, name), "WEBP", quality=84, method=6)
    return name


def 写真を差し替える(target: str, data: bytes) -> None:
    """save_image: 既存の写真を同じ名前で入れ替える。元は admin/backups/_images/ に控える。"""
    from PIL import Image

    core = repo.部品("core")
    img_dir = _img_dir()
    dst = os.path.join(img_dir, os.path.basename(target))
    bak = os.path.join(core.BACKUP_DIR, "_images")
    os.makedirs(bak, exist_ok=True)
    if os.path.exists(dst):
        shutil.copy2(dst, os.path.join(bak, os.path.basename(target)))
    im = Image.open(io.BytesIO(data))
    if im.width > 1600:
        im = im.resize((1600, round(im.height * 1600 / im.width)), Image.LANCZOS)
    ext = target.rsplit(".", 1)[-1].lower()
    if ext == "webp":
        im.convert("RGB" if im.mode in ("CMYK", "P") else im.mode).save(dst, "WEBP", quality=84, method=6)
    elif ext in ("jpg", "jpeg"):
        im.convert("RGB").save(dst, "JPEG", quality=88, optimize=True)
    else:
        im.save(dst)
