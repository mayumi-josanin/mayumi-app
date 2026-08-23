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
from django.views.decorators.csrf import csrf_exempt

from apps.members.models import Member
from apps.records.models import AppSetting
from apps.content.models import (
    CalendarEvent,
    Category,
    Menu,
    News,
    Product,
    PushNotice,
    SupportFaq,
)


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



def _引数(request):
    """GAS と同じ形で引数を受け取る。

    アプリは `?action=... &data={"memberId":"..."}` の形で送ってくる
    （`fetchFromGAS(action, params)` を見ると `data` に JSON を入れている）。
    `?memberId=...` の形でも受けられるようにしておく。
    """
    import json as _json

    生 = request.GET.get("data", "")
    if 生:
        try:
            v = _json.loads(生)
            if isinstance(v, dict):
                return v
        except ValueError:
            pass
    return request.GET.dict()


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



def _画像(値):
    """GAS の parseStoredImageUrls_ と同じ考え方。

    1枚だけの列でも、GAS は配列にして返している。
    空なら空の配列。**None を返さない。**
    """
    v = (値 or "").strip()
    return [v] if v else []


def _メニュー():
    """GAS の getMenus()（9176行）と同じ形で返す。

    **お知らせと除外条件が違う。**メニューは「公開設定が『公開』であること」。
    お知らせは「お知らせ一覧公開」だけを見ていた。**まとめて書かない。**
    """
    from django.utils import timezone
    いま = timezone.now()

    一覧 = []
    for m in Menu.objects.all():
        # GAS（9176行）が外している4つ:
        #   row[5] !== '公開'        … **公開設定が「公開」でなければ外す**
        #   isNoticeListingDeleted_  … お知らせ一覧削除日時
        #   isSoftDeletedByColumns_  … 削除状態・削除日時
        #   !isPublishAtAvailable_   … 公開開始日時がまだ
        if not m.published:
            continue
        if m.notice_delisted_at:
            continue
        if m.deleted or m.deleted_at:
            continue
        if m.publish_at and m.publish_at > いま:
            continue

        画像 = list(m.image_urls or [])
        一覧.append({
            "rowIdx": m.sheet_row,
            "date": _日付(m.registered_on),
            "name": m.name or "",
            # GAS は「画像が無ければ元の列の値をそのまま返す」作り。
            # サーバーでは配列で持っているので、先頭か空文字。
            "imageUrl": 画像[0] if 画像 else "",
            "imageUrls": 画像,
            "description": m.summary or "",
            "reservationStatus": m.booking_status or "",
            "category": m.category or "",
            "updatedAt": _日時(m.updated_at),
            "publishAt": _日時(m.publish_at),
            "noticeStatus": "公開" if m.notice_listed else "非公開",
            "sortOrder": m.sort_key or 0,
        })
    return {"status": "ok", "menus": 一覧}


def _カレンダー():
    """GAS の getCalendarEvents()（**8292行のほう**）と同じ形で返す。

    3344行にも同じ名前の関数があるが、**後から読まれる8292行が勝つ。**
    1つ目を写すと、返す項目が足りなくなる。
    """
    from django.utils import timezone
    いま = timezone.now()

    一覧 = []
    for c in CalendarEvent.objects.all():
        # メニューと同じく「公開設定が『公開』であること」
        if not c.published:
            continue
        if c.notice_delisted_at:
            continue
        if c.deleted or c.deleted_at:
            continue
        if c.publish_at and c.publish_at > いま:
            continue

        画像 = _画像(c.image_url)
        一覧.append({
            "rowIdx": c.sheet_row,
            "date": _日付(c.event_on),
            "title": c.title or "",
            "desc": c.detail or "",
            "color": c.color or "",
            # カレンダーのカテゴリ列は、サーバーでは持っていない。
            # **無いものを作らない。**空文字で返す（GASも空のことが多い）。
            "category": "",
            "image": 画像[0] if 画像 else "",
            "imageUrls": 画像,
            "updatedAt": _日時(c.updated_at),
            "publishAt": _日時(c.publish_at),
            # リンクURL・ボタンテキストはシートに列があるが**どちらも0件**なので
            # 移していない。使われ始めたら移す（2026-08-22 の点検で確認）。
            "linkUrl": "",
            "linkButtonText": "",
            "menuRowIdx": c.menu_row or 0,
            "noticeStatus": "公開" if c.notice_listed else "非公開",
            "sortOrder": c.sort_order or 0,
        })
    return {"status": "ok", "events": 一覧}


def _使い方FAQ():
    """GAS の getSupportFaq()（8994行）と同じ形で返す。

    公開のものだけ（`getSupportFaqEntries_(false)`）。
    並び順は**優先度の高い順、同点はシートの行の順**。
    """
    一覧 = [
        {
            "rowIdx": f.sheet_row,
            "category": f.category or "",
            "question": f.question or "",
            "keywords": f.keywords or "",
            "answer": f.answer or "",
            "priority": f.priority or 0,
            "updatedAt": _日時(f.updated_at),
        }
        # GAS は question と answer が両方ある公開のものだけを出す
        for f in SupportFaq.objects.filter(published=True).exclude(question="").exclude(answer="")
    ]
    一覧.sort(key=lambda x: (-x["priority"], x["rowIdx"]))
    return {"status": "ok", "faqs": 一覧}


def _通知():
    """GAS の getPushNotices()（8158行）と同じ形で返す。

    **並びは日時の新しい順。**GAS は sentAt / scheduledAt / updatedAt の
    どれかを日時とみなして並べている（前のものが空なら次を使う）。
    """
    def 日時の数(p):
        for d in (p.sent_at, p.scheduled_at, p.updated_at):
            if d:
                return int(d.timestamp() * 1000)
        return 0

    一覧 = []
    for p in PushNotice.objects.all():
        if p.deleted or p.deleted_at:
            continue
        一覧.append({
            "rowIdx": p.sheet_row,
            "date": 日時の数(p),
            "sentAt": _日時(p.sent_at),
            "scheduledAt": _日時(p.scheduled_at),
            "updatedAt": _日時(p.updated_at),
            "title": p.title or "",
            "body": p.body or "",
            "targetStatus": p.target_status or "all",
            "targetDetail": p.target_detail or "",
            "recipientCount": p.recipient_count or 0,
            "status": p.status or "",
            "targetPage": p.target_page or "home",
            "previewBody": p.preview_body or "",
            "notificationId": p.notification_id or "",
            "result": p.result or "",
        })
    一覧.sort(key=lambda x: -(x["date"] or 0))
    return {"status": "ok", "notices": 一覧}



def _商品():
    """GAS の getProducts()（3420行）と同じ形で返す。

    **除外条件がメニュー・カレンダーと逆。**

        getMenus    … 公開設定が「公開」**でなければ外す**（空欄は外れる）
        getProducts … 公開設定が「非公開」**なら外す**（**空欄は残る**）

    写し間違えると、出してはいけない商品が出るか、出るべき商品が消える。

    **売切は在庫数ではなく「売切状態」列で決まる。**
    在庫が残っていても院長が手で「売切」にできる。
    この列は 2026-08-17 の移行で見落としていて、8/22 に足した。
    """
    from django.utils import timezone
    いま = timezone.now()

    一覧 = []
    for p in Product.objects.all():
        # GAS（3420行）が外している4つ:
        #   !row[1]                     … 商品名が空
        #   row[5] === '非公開'          … **「非公開」のときだけ外す**
        #   isNoticeListingDeleted_     … お知らせ一覧削除日時
        #   isSoftDeletedByColumns_     … 削除状態・削除日時
        #   !isPublishAtAvailable_      … 公開開始日時がまだ
        if not (p.name or "").strip():
            continue
        # **GAS には無い条件を足さない。**
        # 最初、「空欄も False になってしまうから」と考えて
        # notice_listed で救う条件を書いたが、GAS にそんな条件は無い。
        # そのせいで非公開の「天然だし調味粉」が出てしまい、9件になった
        # （GAS は8件）。**推測で条件を足すと、出してはいけないものが出る。**
        if not p.published:
            continue
        if p.notice_delisted_at:
            continue
        if p.deleted or p.deleted_at:
            continue
        if p.publish_at and p.publish_at > いま:
            continue

        在庫 = p.stock or 0
        閾値 = p.stock_warning or 0
        売切 = (p.sold_out or "").strip() == "売切"

        画像 = _画像(p.icon_url)
        説明画像 = _画像(p.description_image_url)

        一覧.append({
            "category": p.category or "",
            "name": p.name or "",
            "price": p.price or 0,
            # GAS は「画像があればその1枚目、無ければ元の列の値（絵文字）」
            "icon": 画像[0] if 画像 else (p.icon_url or "🌿"),
            "imageUrls": 画像,
            "bg": p.background_color or "c1",
            "description": p.description or "",
            "descriptionImage": 説明画像[0] if 説明画像 else (p.description_image_url or ""),
            "descriptionImageUrls": 説明画像,
            "updatedAt": _日時(p.updated_at),
            "stockQty": 在庫,
            "lowStockThreshold": 閾値,
            "soldOutStatus": "売切" if 売切 else "在庫あり",
            "isSoldOut": 売切,
            # **売切でない かつ 閾値>0 かつ 在庫>0 かつ 在庫<=閾値**
            "isLowStock": (not 売切) and 閾値 > 0 and 在庫 > 0 and 在庫 <= 閾値,
            "publishAt": _日時(p.publish_at),
            "noticeStatus": "公開" if p.notice_listed else "非公開",
        })
    return {"status": "ok", "products": 一覧}



# 端末の記録は8件まで。GAS の MAX_DEVICE_SESSIONS と同じ値。
# **ここを変えると、GASと違う数を返すことになる。**
端末の上限 = 8


def _端末(request):
    """GAS の getUserDevices()（7680行）と同じ形で返す。

    会員IDが要る。無ければエラー（GASと同じ文言）。
    見つからない会員は**エラーではなく空の一覧**を返す（GASもそうしている）。
    """
    会員ID = _引数(request).get("memberId", "")
    会員ID = str(会員ID or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが必要です"}

    m = Member.objects.filter(member_id=会員ID).first()
    if not m:
        return {"status": "ok", "devices": []}

    生 = m.device_sessions or []
    if not isinstance(生, list):
        return {"status": "ok", "devices": []}

    一覧 = []
    for d in 生:
        if not isinstance(d, dict):
            continue
        端末ID = str(d.get("deviceId") or "").strip()
        # **deviceId が無いものは捨てる。**GAS も同じ。
        if not 端末ID:
            continue
        一覧.append({
            "deviceId": 端末ID,
            "label": str(d.get("label") or "").strip(),
            "platform": str(d.get("platform") or "").strip(),
            "appVersion": str(d.get("appVersion") or "").strip(),
            "lastSeenAt": str(d.get("lastSeenAt") or "").strip(),
            # **=== true のときだけ真。**「あれば真」にすると
            # 文字列の "false" が真になってしまう。
            "passcodeEnabled": d.get("passcodeEnabled") is True,
            "pushEnabled": d.get("pushEnabled") is True,
            "current": d.get("current") is True,
        })

    # 最後に使った順。GAS と同じ並び。
    一覧.sort(key=lambda x: x["lastSeenAt"] or "", reverse=True)
    return {"status": "ok", "devices": 一覧[:端末の上限]}


def _注文(request):
    """GAS の getCustomerOrders()（4869行）と同じ形で返す。

    **注文管理シートは0行。**GAS も必ず空を返している。
    サーバーにも注文の表は作っていない（[移行設計] の判断A「作らない」）。

    それでも口だけは用意する。**無いと 501 を返してしまい、
    アプリ側が「サーバーが壊れている」と受け取る恐れがある。**
    """
    会員ID = str(_引数(request).get("memberId", "") or "").strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが必要です"}
    return {"status": "ok", "orders": []}



# GAS の DEFAULT_APP_RUNTIME_CONFIG と同じ既定値。
# 保存されていない項目は、GAS もこれで補っている。
アプリ設定の既定 = {
    "latestAppVersion": "1.1.1",
    "minimumSupportedVersion": "0.0.0",
    "iosStoreUrl": "",
    "updateTitle": "アップデートが必要です",
    "updateMessage": "このアプリを引き続き利用するには、最新版へアップデートしてください。",
    "webBundleVersion": "2026.04.05.61",
}


def _版を比べる(a, b):
    """GAS の compareVersions_ と同じ。a が b より小さければ負を返す。"""
    def 数に(v):
        out = []
        for x in str(v or "").split("."):
            try:
                out.append(int(x))
            except ValueError:
                out.append(0)
        return out

    x, y = 数に(a), 数に(b)
    for i in range(max(len(x), len(y))):
        p = x[i] if i < len(x) else 0
        q = y[i] if i < len(y) else 0
        if p != q:
            return p - q
    return 0


def _アプリ設定():
    """GAS の getAppRuntimeConfig()（466行）と同じ形で返す。

    **保存された値をそのまま返さない。**GAS は `sanitizeAppRuntimeConfig_` を
    通しており、そこで3つのことをしている。同じことをする。

    1. 空の項目は既定値で補う
    2. **`webBundleVersion` は保存値を無視して既定値で固定**
    3. 最新版が最低対応版より古ければ、最低対応版に揃える
    4. `iosStoreUrl` が http(s) で始まらなければ空にする

    2 は GAS のコメントにこう書いてある。

        // 現在配布中のネイティブ版との整合を優先し、返却値は固定で揃える

    実際、保存値は `2026.04.06.63` だが GAS は `2026.04.05.61` を返している
    （2026-08-22 に確認）。**院長の判断で、この動きをそのまま写す。**
    ここを「保存値を返す」に変えると、**お客様に更新案内が出はじめる恐れがある。**
    """
    行 = AppSetting.objects.filter(key="APP_RUNTIME_CONFIG").first()
    保存 = (行.value if 行 and isinstance(行.value, dict) else {}) or {}

    設定 = {}
    for k in ("latestAppVersion", "minimumSupportedVersion", "updateTitle", "updateMessage"):
        v = str(保存.get(k) or "").strip()
        設定[k] = v or アプリ設定の既定[k]

    設定["iosStoreUrl"] = str(保存.get("iosStoreUrl") or アプリ設定の既定["iosStoreUrl"]).strip()

    # **保存値を使わない。**GAS と同じく固定。
    設定["webBundleVersion"] = アプリ設定の既定["webBundleVersion"]

    if _版を比べる(設定["latestAppVersion"], 設定["minimumSupportedVersion"]) < 0:
        設定["latestAppVersion"] = 設定["minimumSupportedVersion"]

    if 設定["iosStoreUrl"] and not 設定["iosStoreUrl"].lower().startswith(("http://", "https://")):
        設定["iosStoreUrl"] = ""

    # GAS の返す並びに合わせる（JSONの順は見た目だけの話だが、
    # 突き合わせのときに読みやすい）
    順 = ["latestAppVersion", "minimumSupportedVersion", "iosStoreUrl",
          "updateTitle", "updateMessage", "webBundleVersion"]
    return {"status": "ok", "config": {k: 設定[k] for k in 順}}


def _ガチャ設定():
    """GAS の getRewardGachaConfig()（644行）と同じ形で返す。

    月ごとの景品表。中身の形は GAS が保存したまま。
    **手を加えない。**確率や文言を勝手に整えると、当たり方が変わる。
    """
    行 = AppSetting.objects.filter(key="REWARD_GACHA_CONFIG").first()
    値 = (行.value if 行 and isinstance(行.value, dict) else None) or {"monthlyPrizes": []}
    if "monthlyPrizes" not in 値:
        値 = {"monthlyPrizes": []}
    return {"status": "ok", "config": 値}


# action の名前 → 返す中身を作る関数
_できること = {
    "getNews": _お知らせ,
    "getMenus": _メニュー,
    "getCalendar": _カレンダー,
    "getSupportFaq": _使い方FAQ,
    "getPushNotices": _通知,
    "getProducts": _商品,
    # 会員IDが要るもの。request を受け取る。
    "getUserDevices": _端末,
    "getCustomerOrders": _注文,
    "getAppRuntimeConfig": _アプリ設定,
    "getRewardGachaConfig": _ガチャ設定,
}


@csrf_exempt
def 窓口(request):
    """GAS の doGet / doPost と同じ入口。

    読み取りは `?action=...`、書き込みは POST の `{"type": "..."}`。
    **GAS がその形なので、そのまま合わせる。**

    まだ作っていない action は、**素直に「まだありません」と答える。**
    黙って空を返すと、切り替えたときに「データが消えた」ように見える。

    `csrf_exempt` を付けるのは、アプリが別の場所（GitHub Pages）から
    呼ぶため。**代わりに合鍵（API_KEY）で守っている。**
    """
    断り = _合鍵を確かめる(request)
    if 断り:
        return 断り

    # 書き込み。GAS の doPost にあたる。
    if request.method == "POST":
        from .writes import 受け取る

        書く, 中 = 受け取る(request)
        if not 書く:
            # 中 にはエラーの中身が入っている。
            # **GAS はエラーでも 200 で返す。**同じにする。
            return JsonResponse(中, json_dumps_params={"ensure_ascii": False})
        return JsonResponse(書く(中), json_dumps_params={"ensure_ascii": False})

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

    # 引数が要るものだけ request を渡す。
    # **引数の有無で呼び方を変える。**全部に request を渡す作りにすると、
    # 使わない関数まで引数を持たされて、読みにくくなる。
    import inspect

    if inspect.signature(作る).parameters:
        中身 = 作る(request)
    else:
        中身 = 作る()

    # 会員IDが無いときなど、GAS はエラーでも 200 で返している。
    # **同じにする。**アプリは status を見て判断している。
    return JsonResponse(中身, json_dumps_params={"ensure_ascii": False})
