"""お客様の画面（確認用）。お客様が見ている3つのアプリを、管理画面の枠の中でそのまま確かめる。

院長の依頼（2026-09-18）:「お客様アプリがあるやつ（ビジリスアプリ・まゆみ助産院お客様アプリ・予約システム）を
develop 環境で確認して見れるようにできますか」。前の日に作ったビジリスだけの画面（apps/bijiris/views_preview.py）を
ここへ引っ越して、3つを同じ作りの1か所にまとめた。

どれも「2つを見比べる」だけの画面:

  いま作っている方（develop） … サーバーの手元にある develop の写し（settings.PREVIEW_APP_DIR）のファイルを配る。
                                予約システムだけは Django なので、確認用の住所をそのまま枠に出す。
  お客様に出ている方（main）   … 公開されている本物（GitHub Pages・予約システムの本番）を枠に読み込む。

作りは公式サイトの「プレビュー」（apps/hp/views.py）と同じに揃えてある。
  - ファイルを配る view には @xframe_options_sameorigin（既定の DENY だと枠の中で「接続が拒否されました」になる）
  - 出すのは決めた置き場の下だけ。".." を含む道は弾く

**配ってよいものは「通してよい一覧」で決める。**まゆみ助産院アプリは写しの根っこをそのまま配るので、
拒否の一覧（server を出さない…）にすると、ものが増えたときに漏れる。許すものだけを並べておけば、
知らないものは既定で通らない。**server/（この Django のコード・.env・backups）は絶対に出さない。**

読むだけの画面なので、権限は他の管理画面と同じでまゆみだけ（owner_required）。
"""

import mimetypes
import os

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import render
from django.urls import reverse
from django.views.decorators.clickjacking import xframe_options_sameorigin

from apps.manage.permissions import owner_required

# 枠の中の幅。スマホは院長の依頼どおり 375px、パソコンは枠いっぱい（0）
DEVICES = [("スマホ", 375), ("パソコン", 0)]

# まゆみ助産院アプリで**通してよいもの**（写しの根っこにある、お客様アプリの部品だけ）。
# ここに無いものは配らない。server（この Django のコード・.env・backups）・.git・docs・gas・node_modules は
# 並べていないので通らない。お客様アプリに部品が増えたときは、ここに足す。
アプリの通してよい一覧 = {
    # お客様アプリ本体
    "index.html", "app.js", "style.css", "sw.js", "manifest.json", "OneSignalSDKWorker.js",
    "stamp-launch.html",
    # 画像
    "icon.png", "logo-mayumi.png", "instr_ios_1.png", "instr_ios_2.png", "instr_ios_3.png",
    "icons", "img", "assets",
    # アプリから開く別ページ
    "hajimekata", "start", "manuals",
}

# 3つのタブ。並びはこの順で左メニューに出る（apps/manage/navigation.py）
種類たち = {
    "app": {
        "label": "まゆみ助産院アプリ",
        # お客様に出ている方（main）。GitHub Pages に出している本物
        "公開先": "https://mayumi-josanin.github.io/mayumi-app/",
        # いま作っている方（develop）。写しの根っこがそのままお客様アプリ
        "下の道": "",
        "通してよい一覧": アプリの通してよい一覧,
    },
    "bijiris": {
        "label": "ビジリス",
        "公開先": "https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/",
        "下の道": os.path.join("bijiris", "customer-app"),
        # customer-app の下はまるごとお客様アプリなので、一覧で絞る必要がない
        "通してよい一覧": None,
    },
    "reserve": {
        "label": "予約システム",
        # 予約システムは Django（mayumi-reserve）。ファイルではないので配らず、住所をそのまま枠に出す
        "住所だけ": True,
    },
}


def 置き場(kind: str) -> str:
    """いま作っている方（develop の写し）の場所。

    ビジリスは、置き方が2通りある。
      ・mayumi-app … お客様アプリの写しの下（bijiris/customer-app/）
      ・まとめた場所 … apps/customer と apps/bijiris が並び。お客様アプリの下に無い
    後者では PREVIEW_BIJIRIS_DIR でビジリスの場所を教える（決めてなければ今までどおり）。
    """
    別置き = settings.PREVIEW_BIJIRIS_DIR if kind == "bijiris" else ""
    if 別置き:
        return 別置き
    return os.path.join(settings.PREVIEW_APP_DIR or "", 種類たち[kind]["下の道"])


def ファイルがあるか(kind: str) -> bool:
    return os.path.isfile(os.path.join(置き場(kind), "index.html"))


def _画面(request, kind: str):
    設定 = 種類たち[kind]
    住所だけ = 設定.get("住所だけ", False)
    # 予約システムは Django が動いているので「手元のファイル」は無い。いつでも住所を出す
    手元あり = True if 住所だけ else ファイルがあるか(kind)
    # どちらを見るか。既定は「いま作っている方」。手元にファイルが無いときだけ公開の方を先に出す
    which = request.GET.get("which") or ("dev" if 手元あり else "live")
    if which not in ("dev", "live"):
        which = "dev"
    width = request.GET.get("w", "375")
    if width not in [str(w) for _, w in DEVICES]:
        width = "375"
    if 住所だけ:
        dev_url = settings.PREVIEW_RESERVE_DEV_URL
        公開先 = settings.PREVIEW_RESERVE_LIVE_URL
    else:
        dev_url = reverse("manage:preview:files", kwargs={"kind": kind, "path": ""})
        公開先 = 設定["公開先"]
    return render(request, "preview/preview.html", {
        # 枠に出す住所。「いま作っている方」か「お客様に出ている方」か
        "src": 公開先 if which == "live" else dev_url,
        "kind": kind, "label": 設定["label"], "which": which, "width": width, "devices": DEVICES,
        "local_ready": 手元あり, "public_url": 公開先, "dev_url": dev_url,
        "住所だけ": 住所だけ, "dir": "" if 住所だけ else 置き場(kind),
    })


@owner_required
def app(request):
    """まゆみ助産院アプリ（お客様アプリ）。"""
    return _画面(request, "app")


@owner_required
def bijiris(request):
    """ビジリス（アンケートアプリ）。"""
    return _画面(request, "bijiris")


@owner_required
def reserve(request):
    """予約システム。ファイルではなく、確認用の住所をそのまま枠に出す。"""
    return _画面(request, "reserve")


def _配る(基点: str, パス: str, 通してよい一覧):
    """基点の中の、通してよいものだけを返す。外へは出さない。"""
    パス = (パス or "").lstrip("/")
    if パス == "" or パス.endswith("/"):
        パス += "index.html"
    区切り = パス.split("/")
    if ".." in 区切り:
        raise Http404()
    # **許す一覧で守る。**拒否の一覧だと、写しにものが増えたときに漏れる
    if 通してよい一覧 is not None and 区切り[0] not in 通してよい一覧:
        raise Http404()
    実体 = os.path.abspath(os.path.join(基点, パス))
    if not 実体.startswith(os.path.abspath(基点) + os.sep) or not os.path.isfile(実体):
        raise Http404()
    種類, _ = mimetypes.guess_type(実体)
    return FileResponse(open(実体, "rb"), content_type=種類 or "application/octet-stream")


@owner_required
@xframe_options_sameorigin  # 管理画面の枠（iframe）に出すため。既定の DENY だと Chrome が「接続が拒否されました」と出す
def files(request, kind="app", path=""):
    """いま作っている方（develop の写し）のファイルを配る。"""
    設定 = 種類たち.get(kind)
    if 設定 is None or 設定.get("住所だけ", False):
        raise Http404()
    return _配る(置き場(kind), path, 設定["通してよい一覧"])
