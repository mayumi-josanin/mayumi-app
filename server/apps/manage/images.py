"""お知らせに付ける画像を、このサーバーに置く。

GAS の管理アプリは Google Drive に上げて thumbnail の URL を持っていた。
ここでは `media/news/` に置き、Funnel（mayumi-api）の `/media/news/<名前>` で配る。
お客様アプリは URL を受け取って表示するだけなので、置き場が変わっても困らない。

**名前は乱数。**推測で他の画像を見に行けないようにするのと、
同じ名前で上書きされて古いお知らせの画像が変わるのを防ぐため。
"""

import io
import uuid
from pathlib import Path

from django.conf import settings

MAX_EDGE = 1600
JPEG_QUALITY = 85


def 保存する(uploaded) -> str:
    """アップロードされた1枚を縮めて保存し、公開URLを返す。"""
    from PIL import Image, ImageOps

    img = Image.open(uploaded)
    # スマートフォンの写真は EXIF で向きを持つ。そのまま保存すると横倒しになる。
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.thumbnail((MAX_EDGE, MAX_EDGE))

    置き場 = Path(settings.MEDIA_ROOT) / "news"
    置き場.mkdir(parents=True, exist_ok=True)
    名前 = f"{uuid.uuid4().hex}.jpg"
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
    (置き場 / 名前).write_bytes(buf.getvalue())
    return f"{settings.PUBLIC_BASE_URL}/media/news/{名前}"
