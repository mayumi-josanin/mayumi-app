"""ビジリス管理「お客様の画面」。お客様が見ているアンケートアプリを、管理画面の枠の中で確かめる。

院長の依頼（2026-09-17）:「ビジリス管理でお客様のアプリ画面を確認できるようにしたい。develop と main で」
2つを見比べられるようにする。

  いま作っている方（develop）  … サーバーの手元にある bijiris/customer-app/ のファイルをそのまま出す。
                                 置き場所は settings.BIJIRIS_DIR（docker-compose の manage に /bijiris で載せている）。
  お客様に出ている方（main）   … GitHub Pages に出ている本物（公開先）を枠の中に読み込む。

作りは公式サイトの「プレビュー」（apps/hp/views.py の hp_preview_file / hp_site_file）と同じに揃えてある。
  - ファイルを配る view には @xframe_options_sameorigin を付ける（既定の DENY だと枠の中で「接続が拒否されました」になる）
  - 出すのは customer-app/ の下だけ。".." を含む道は弾く（hp の _配る と同じ守り）

読むだけの画面なので gate（サーバーが正かどうか）は関係しない。権限は他のビジリス管理と同じでまゆみだけ。
"""

import mimetypes
import os

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import render
from django.views.decorators.clickjacking import xframe_options_sameorigin

from apps.manage.permissions import owner_required

# お客様に出ている方（main）の住所。GitHub Pages に出しているビジリスのお客様アプリ。
公開先 = "https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/"

# 枠の中の幅。スマホは院長の依頼どおり 375px、パソコンは枠いっぱい（0）
DEVICES = [("スマホ", 375), ("パソコン", 0)]


def 置き場():
    """手元のお客様アプリ（bijiris/customer-app）の場所。"""
    return os.path.join(settings.BIJIRIS_DIR or "", "customer-app")


def ファイルがあるか() -> bool:
    return os.path.isfile(os.path.join(置き場(), "index.html"))


@owner_required
def preview(request):
    """「お客様の画面」タブ。いま作っている方／お客様に出ている方 を切り替えて枠の中に出す。"""
    # どちらを見るか。既定は「いま作っている方」。手元にファイルが無いときだけ公開の方を先に出す
    手元あり = ファイルがあるか()
    which = request.GET.get("which") or ("dev" if 手元あり else "live")
    if which not in ("dev", "live"):
        which = "dev"
    width = request.GET.get("w", "375")
    if width not in [str(w) for _, w in DEVICES]:
        width = "375"
    return render(request, "bijiris/preview.html", {
        "which": which, "width": width, "devices": DEVICES,
        "local_ready": 手元あり, "public_url": 公開先, "dir": 置き場(),
    })


def _配る(基点: str, パス: str):
    """基点（customer-app）の中のファイルだけを返す。外へは出さない。"""
    パス = (パス or "").lstrip("/")
    if パス == "" or パス.endswith("/"):
        パス += "index.html"
    if ".." in パス.split("/"):
        raise Http404()
    実体 = os.path.abspath(os.path.join(基点, パス))
    if not 実体.startswith(os.path.abspath(基点) + os.sep) or not os.path.isfile(実体):
        raise Http404()
    種類, _ = mimetypes.guess_type(実体)
    return FileResponse(open(実体, "rb"), content_type=種類 or "application/octet-stream")


@owner_required
@xframe_options_sameorigin  # 管理画面の枠（iframe）に出すため。既定の DENY だと Chrome が「接続が拒否されました」と出す
def preview_file(request, path=""):
    """いま作っている方（手元の bijiris/customer-app）のファイルを配る。"""
    return _配る(置き場(), path)
