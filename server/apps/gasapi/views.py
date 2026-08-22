"""GAS と同じ答えを返す窓口。

**返す形を GAS に合わせることが、この窓口のすべて。**
1項目でも違うと、お客様の画面が崩れる。GAS の実物を読んで、
同じ名前・同じ形・同じ並びで返す。

いまは**読むだけ。**書き込みは受け付けない。
スプレッドシートが正のままなので、ここで書けてしまうと
「どちらが正しいか」が崩れる。

## 落とし穴（2026-08-22 に踏んだもの）

**GAS は同じ名前の関数が二重に定義されていることがある。**
後から読んだほうで上書きされるので、**前のほうは動いていない。**
写すときは `grep -n "^function 名前"` で行番号を出し、**大きいほうを読む。**

    getBlogNews        3279 / **8668**  ← こちらが動く
    getCalendarEvents  3344 / **8292**
    getAdminBlogs      4938 / **8737**
    getAdminCalendar   6558 / **8348**
    handleUpdateOrder  3201 / **4829**
"""

import hmac

from django.conf import settings
from django.http import JsonResponse

from apps.content.models import Category, News


def _合鍵を確かめる(request):
    """合鍵がなければ断る。無ければ None を返す（＝通してよい）。

    measurements と同じ作りにしてある。かかった時間から鍵を
    言い当てられないよう compare_digest を使う。
    """
    expected = settings.API_KEY
    if not expected:
        return JsonResponse(
            {"status": "error", "message": "APIキーが未設定です。"}, status=503
        )
    given = request.headers.get("X-Api-Key", "")
    if not hmac.compare_digest(given.encode("utf-8", "ignore"), expected.encode("utf-8")):
        return JsonResponse({"status": "error", "message": "権限がありません。"}, status=403)
    return None


def _日時(値):
    """GAS の formatMaybeDateTime_ と同じ形にする。

    GAS:  Date なら "yyyy-MM-dd'T'HH:mm:ssXXX"、それ以外は文字列のまま
    空なら空文字。**None を返さない。**アプリが文字列として扱うため。
    """
    if not 値:
        return ""
    return 値.astimezone().strftime("%Y-%m-%dT%H:%M:%S%z").replace("+0900", "+09:00")


def _日付(値):
    """投稿日。GAS 側は Date なら日時の形、文字列ならそのまま返している。

    サーバーでは DateField（時刻を持たない）なので、
    **GAS が返していた形に合わせて 00:00:00 を足す。**
    """
    if not 値:
        return ""
    return 値.strftime("%Y-%m-%dT00:00:00+09:00")


def _お知らせ():
    """GAS の getBlogNews()（8668行のほう）と同じ形で返す。

    GAS が出しているもの（実物から書き写した）:

        rowIdx / date / title / category / type / icon / body
        image / imageUrl / imageUrls / updatedAt
        linkUrl / linkButtonText / publishAt / noticeStatus / sortOrder

    **image と imageUrl は同じ値。**古い呼び名を残したまま新しい名前を
    足したもの。片方だけにするとアプリが崩れる恐れがあるので両方返す。
    """
    # カテゴリの種別（お知らせ / ブログ）を先に引く。GAS も同じことをしている。
    種別 = {}
    for c in Category.objects.all():
        名 = (c.name or "").strip()
        if 名:
            種別[名] = "お知らせ" if c.kind == "お知らせ" else "ブログ"

    from django.utils import timezone
    いま = timezone.now()

    一覧 = []
    for n in News.objects.all():
        # GAS（8668行）が外している4つ。**この順・この条件をそのまま写す。**
        #
        #   isNoticeListingDeleted_        … お知らせ一覧削除日時が入っている
        #   isSoftDeletedByColumns_        … 削除状態 or 削除日時が入っている
        #   normalizePublishVisibilityStatus_ … お知らせ一覧公開が「非公開」
        #   !isPublishAtAvailable_          … 公開開始日時がまだ来ていない
        #
        # **「公開設定」（published）は見ていない。**GAS が見るのは
        # 「お知らせ一覧公開」のほうだけ。ここを足すと、GASでは出ているものが
        # サーバーでは消える。最初これを入れてしまい、突き合わせで気づいた。
        if n.notice_delisted_at:
            continue
        if n.deleted or n.deleted_at:
            continue
        if not n.notice_listed:
            continue
        if n.publish_at and n.publish_at > いま:
            continue

        category = n.category or "お知らせ"
        画像 = [u for u in [(n.image_url or "").strip()] if u]
        画像1 = 画像[0] if 画像 else ""

        一覧.append({
            "rowIdx": n.sheet_row,
            "date": _日付(n.posted_on),
            "title": n.title or "",
            "category": category,
            "type": 種別.get(category)
                    or ("お知らせ" if category in ("お知らせ", "休診情報") else "ブログ"),
            "icon": n.icon or "📢",
            "body": n.body or "",
            "image": 画像1,
            "imageUrl": 画像1,
            "imageUrls": 画像,
            "updatedAt": _日時(n.updated_at),
            "linkUrl": (n.link_url or "").strip(),
            "linkButtonText": (n.link_label or n.button_text or "").strip(),
            "publishAt": _日時(n.publish_at),
            "noticeStatus": "公開" if n.notice_listed else "非公開",
            "sortOrder": n.sort_order or 0,
        })

    # GAS は categories も一緒に返している。アプリが同じ呼び出しで使うため。
    カテゴリ = [
        {"name": c.name, "type": c.kind or "ブログ", "rowIdx": c.sheet_row}
        for c in Category.objects.all()
    ]
    return {"status": "ok", "news": 一覧, "categories": カテゴリ}


# action の名前 → 返す中身を作る関数
_できること = {
    "getNews": _お知らせ,
}


def 窓口(request):
    """GAS の doGet と同じ入口。?action=... で振り分ける。

    まだ作っていない action は、**素直に「まだありません」と答える。**
    黙って空を返すと、切り替えたときに「データが消えた」ように見える。
    """
    断り = _合鍵を確かめる(request)
    if 断り:
        return 断り

    action = request.GET.get("action", "").strip()
    if not action:
        return JsonResponse(
            {"status": "error", "message": "action を指定してください。"}, status=400
        )

    作る = _できること.get(action)
    if not 作る:
        return JsonResponse({
            "status": "error",
            "message": f"'{action}' はまだサーバー側にありません。GASをお使いください。",
            "notImplemented": True,
        }, status=501)

    return JsonResponse(作る(), json_dumps_params={"ensure_ascii": False})
