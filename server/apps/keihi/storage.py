"""レシートの写真を置く。

**公開の /media/<置き場>/ には置かない。** あちらはお客様アプリと同じサーバーにも生えていて、
リンクを知っていれば誰でも見られる。レシートは金額・店名・日付が写った帳簿の証憑なので、
管理画面にログインした人だけが見られる専用の配り口（views.receipt_image）から出す。

名前は乱数。推測で他のレシートを見に行けないようにするため。
"""

import uuid
from pathlib import Path

from django.conf import settings

# 読み取りに使うので、文字がつぶれない程度には残す。長辺 2000px あればレシートの細かい数字も読める
MAX_EDGE = 2000
JPEG_QUALITY = 85
置き場 = "receipts"


def 置き場所() -> Path:
    path = Path(settings.MEDIA_ROOT) / 置き場
    path.mkdir(parents=True, exist_ok=True)
    return path


def 保存する(uploaded) -> str:
    """アップロードされた1枚を縮めて保存し、ファイル名を返す。"""
    from PIL import Image, ImageOps

    img = Image.open(uploaded)
    # スマートフォンの写真は EXIF で向きを持つ。そのまま保存すると横倒しになる
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.thumbnail((MAX_EDGE, MAX_EDGE))

    名前 = f"{uuid.uuid4().hex}.jpg"
    img.save(置き場所() / 名前, "JPEG", quality=JPEG_QUALITY, optimize=True)
    return 名前


def 読み出す(名前: str) -> bytes | None:
    # 「../」などで外へ出ようとする名前は弾く
    if not 名前 or "/" in 名前 or "\\" in 名前 or ".." in 名前:
        return None
    path = 置き場所() / 名前
    if not path.is_file():
        return None
    return path.read_bytes()


def 消す(名前: str) -> None:
    if not 名前 or "/" in 名前 or "\\" in 名前 or ".." in 名前:
        return
    path = 置き場所() / 名前
    if path.is_file():
        path.unlink()
