"""お客様が上げた写真を配る窓口。

**なぜ Django で配るのか。**

このサーバーは静的ファイルを WhiteNoise で配っている。ところが WhiteNoise は
**起動したときに一度だけ**フォルダを走査する。あとから足したファイルは
存在しないものとして 404 になる。

プロフィール写真は「上げた直後に見える」ことに意味がある。上げてから
サーバーを作り直すまで見えない、では使いものにならない。だから専用に配る。

**いずれ、アプリ経由でしか見えないようにする。**（院長のご希望・2026-08）
いまは GAS と同じ「リンクを知っていれば見える」に揃えてある。切り替えで
見え方まで変えると、何が原因で写真が出ないのか切り分けられなくなるため。
ここを1か所にしておけば、あとから合鍵を求める形に変えられる。
"""

import mimetypes
from pathlib import Path

from django.conf import settings
from django.http import FileResponse, Http404


def 写真(request, 名前):
    """`/media/avatars/<名前>` を配る。

    **名前は英数字と一部の記号だけを通す。**`..` を含む名前で
    サーバーの別の場所を読み出されるのを防ぐ。
    """
    if not 名前 or "/" in 名前 or "\\" in 名前 or ".." in 名前:
        raise Http404

    置き場 = Path(settings.MEDIA_ROOT) / "avatars"
    道 = (置き場 / 名前).resolve()

    # **解決したあとの道が、置き場の中にあることを確かめる。**
    # 名前の見た目だけで判断すると、記号の組み合わせで抜けられることがある。
    try:
        道.relative_to(置き場.resolve())
    except ValueError:
        raise Http404

    if not 道.is_file():
        raise Http404

    種類 = mimetypes.guess_type(str(道))[0] or "application/octet-stream"
    応答 = FileResponse(open(道, "rb"), content_type=種類)
    # 同じ名前のファイルは中身が変わらない（名前に乱数を含めているため）。
    # 何度も取りに来させる理由がない。
    応答["Cache-Control"] = "public, max-age=31536000, immutable"
    return 応答
