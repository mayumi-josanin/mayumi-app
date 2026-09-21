"""QRコード案内。**中身は旧管理アプリ（admin/index.html の #page-qrcode）と同じ。**

出すものは3つ。

  アプリ配布用QRコード    お客様アプリの住所のQR。初回登録やホーム画面追加のご案内用。画像で保存できる
  来院スタンプ用QRコード  院内に貼る、アプリを開くだけのQR（中継ページの住所）
  配布用資料              印刷してお渡しする紙。旧アプリと同じ内容・同じ文言

## QRの絵は、このサーバーが描く

旧アプリは外の絵の作り手（api.qrserver.com）に描いてもらっていた。これだと
**院内から外へ出られないときにQRが出ない**うえ、お客様アプリの住所を毎回外へ渡すことになる。
ここでは segno（QRの絵を描くだけの小さな部品）を使い、**このサーバーの中で描く**。
絵の道は `/manage/qrcode/image/?kind=app`（スタンプ用は `kind=stamp`）。

## 配布用資料は別の窓を開かない

旧アプリは `window.open` で新しい窓を作り、そこへ資料を書き込んで印刷していた。
新しい窓はブラウザに止められることがあり、そうなると何も起きない（押しても無反応に見える）。
ここでは資料を同じ画面の中に置き、**印刷のときだけ資料だけが出る**ようにした（qrcode.html の印刷用の見た目）。

## 住所の決め方

お客様アプリの住所は、設定の置き場（records.AppSetting）の鍵 `app_public_url` に入れる。
画面はシステム管理の中の「お客様アプリの住所」で直す。決まっていないときは、この画面に
その旨とシステム管理への行き先を出す。
"""

import io
from urllib.parse import quote, urljoin

import segno
from django.http import Http404, HttpResponse
from django.shortcuts import render

from apps.records.models import AppSetting

from .permissions import owner_required

鍵 = "app_public_url"

# お客様アプリの住所の既定。records の移行 0011 でこの値を入れてある
既定の住所 = "https://mayumi-josanin.github.io/mayumi-app/"

# 来院スタンプ用QRが指す中継ページ。旧アプリの buildStampLaunchUrl と同じ
スタンプの中継ページ = "stamp-launch.html?action=add_stamp"


# ---- 住所 ----

def アプリの住所() -> str:
    s = AppSetting.objects.filter(pk=鍵).first()
    return str((s.value or {}).get("url") or "").strip() if s else ""


def 住所を決める(住所: str) -> None:
    AppSetting.objects.update_or_create(
        pk=鍵, defaults={"value": {"url": str(住所 or "").strip()},
                        "note": "お客様アプリの住所（QRコード案内で使う）"})


def スタンプの住所(アプリ: str) -> str:
    """中継ページの住所。アプリの住所が入っていなければ空。

    旧アプリは `new URL('stamp-launch.html?action=add_stamp', base)` で組み立てていた。
    形になっていない文字を入れると向こうは空を返したので、ここも同じく空を返す。
    """
    base = str(アプリ or "").strip()
    if not base.startswith(("http://", "https://")):
        return ""
    return urljoin(base, スタンプの中継ページ)


# ---- 画面 ----

@owner_required
def qrcode_view(request):
    住所 = アプリの住所()
    return render(request, "manage/qrcode.html", {
        "app_url": 住所,
        "stamp_url": スタンプの住所(住所),
    })


@owner_required
def qrcode_image(request):
    """QRの絵。`kind=app`（お客様アプリ）か `kind=stamp`（来院スタンプ）。

    `format=svg` と書けば線の絵（大きく印刷しても粗くならない）。既定は png。
    `download=1` を付けると、開かずに保存になる。
    """
    種類 = (request.GET.get("kind") or "app").strip()
    if 種類 not in ("app", "stamp"):
        raise Http404("知らない種類です")
    住所 = アプリの住所()
    値 = スタンプの住所(住所) if 種類 == "stamp" else 住所
    if not 値:
        # 住所が決まっていないときは描かない。空のQRを出すと、読んでも何も起きない紙が出回る
        raise Http404("先にお客様アプリの住所を決めてください")

    形 = (request.GET.get("format") or "png").strip().lower()
    形 = "svg" if 形 == "svg" else "png"
    qr = segno.make(値, error="m")
    # 旧アプリは幅を px で指っていた（300 と 400）。segno は「1マスを何 px にするか」なので、
    # 頼まれた幅に近づくようにマスの大きさを決める
    幅 = max(120, min(1200, _数(request.GET.get("size"), 300)))
    枠 = 2
    マス = max(1, round(幅 / qr.symbol_size(scale=1, border=枠)[0]))

    buf = io.BytesIO()
    qr.save(buf, kind=形, scale=マス, border=枠)
    res = HttpResponse(buf.getvalue(), content_type="image/svg+xml" if 形 == "svg" else "image/png")
    if request.GET.get("download"):
        名 = ("来院スタンプ用QR" if 種類 == "stamp" else "アプリ配布用QR") + "." + 形
        res["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(名)}"
    return res


def _数(v, 既定: int) -> int:
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return 既定
