"""お知らせの表を、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**管理アプリは GAS を見たまま、GAS の中身だけが
シートからここへ変わる。両方のアプリが GAS を見ている限り、正は常に1つ。
（[窓口はGASのまま](docs/design/窓口はGASのまま.md)）

## 通知はここでは送らない

GAS の `handleAddBlog` は、書き込んだあとに OneSignal で通知を送っている。
**通知はGASに残す。**ここが担うのは保存だけ。分けておかないと、
「保存はできたが通知が飛ばない」ときにどちらの落ち度か分からなくなる。

## 行番号を鍵として使い続ける

管理アプリは `rowIdx` でお知らせを指している。**そこを変えない。**
シートへの書き込みが止まっても、`sheet_row` を鍵として持ち続ける。
変えると管理アプリの改修が要り、切り替えが「1つの表を移す」で収まらなくなる。

## 消さずに印を付ける

GAS も論理削除（削除状態の列に印を付ける）。ここも同じにする。
過去の掲載についてのお問い合わせは、実際に来る。
"""

from django.db import transaction
from django.utils import timezone

from apps.content.models import News


def _文(v):
    return "" if v is None else str(v)


def _公開か(値, 既定=True):
    """GAS の '公開' / '非公開' を真偽に直す。"""
    文 = _文(値).strip()
    if not 文:
        return 既定
    return 文 != "非公開"


def _状態の字(公開):
    return "公開" if 公開 else "非公開"


def _日時(v):
    from django.utils.dateparse import parse_datetime

    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_datetime(文)
    if d and timezone.is_naive(d):
        d = timezone.make_aware(d)
    return d


def _日付(v):
    """`2026/9/5` も `2026-09-05` も受ける。シートの書き方が揺れているため。"""
    from django.utils.dateparse import parse_date

    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_date(文.replace("/", "-"))
    if d:
        return d
    try:
        年, 月, 日 = [int(x) for x in 文.replace("-", "/").split("/")[:3]]
        import datetime

        return datetime.date(年, 月, 日)
    except (ValueError, TypeError):
        return None


def _一件(n):
    """GAS の getAdminBlogs が返す1件と同じ形。**項目名を変えない。**"""
    return {
        "rowIdx": n.sheet_row,
        "date": n.posted_on.strftime("%Y-%m-%d") if n.posted_on else "",
        "title": n.title or "",
        "category": n.category or "お知らせ",
        "icon": n.icon or "📢",
        "body": n.body or "",
        "status": _状態の字(n.published),
        "imageUrl": n.image_url or "",
        "publishAt": n.publish_at.strftime("%Y-%m-%dT%H:%M:%S+09:00") if n.publish_at else "",
        "updatedAt": n.updated_at.strftime("%Y-%m-%dT%H:%M:%S+09:00") if n.updated_at else "",
        "noticeStatus": _状態の字(n.notice_listed),
    }


# ── 読む ───────────────────────────────────────────────

def 一覧():
    """GAS の getAdminBlogs と同じ。**消した分は返さない。**

    GAS は最後に `blogs.reverse()` している。シートは古い順に並ぶので、
    新しいものが先頭に来る。ここでは並び順を明示して同じにする。
    """
    件 = News.objects.filter(deleted=False).exclude(title="").order_by("-sheet_row")
    return {"status": "ok", "blogs": [_一件(n) for n in 件]}


# ── 書く ───────────────────────────────────────────────

def 足す(d):
    """GAS の handleAddBlog。**新しい行番号は、いまの最大＋1。**

    シートの appendRow と同じ振る舞いにする。番号が飛んでも構わないが、
    **既にある番号と重なってはいけない。**重なると管理アプリが別の記事を指す。
    """
    題 = _文(d.get("title")).strip()
    if not 題:
        return {"status": "error", "message": "タイトルが必要です"}

    with transaction.atomic():
        最大 = News.objects.select_for_update().order_by("-sheet_row").values_list(
            "sheet_row", flat=True).first() or 1
        n = News.objects.create(
            sheet_row=最大 + 1,
            posted_on=_日付(d.get("date")) or timezone.localdate(),
            title=題,
            category=_文(d.get("category")).strip() or "お知らせ",
            icon=_文(d.get("icon")).strip() or "📢",
            body=_文(d.get("body")),
            image_url=_画像(d),
            link_url=_文(d.get("linkUrl")).strip(),
            button_text=_文(d.get("linkButtonText")).strip(),
            published=_公開か(d.get("status")),
            notice_listed=_公開か(d.get("noticeStatus") or d.get("status")),
            publish_at=_日時(d.get("publishAt")),
            updated_at=timezone.now(),
        )
    return {"status": "ok", "rowIdx": n.sheet_row}


def _画像(d):
    """GAS の serializeStoredImageUrls_ と同じ考え方。複数なら改行でつなぐ。"""
    並び = d.get("imageUrls")
    if isinstance(並び, list) and 並び:
        return "\n".join(_文(x).strip() for x in 並び if _文(x).strip())
    return _文(d.get("image")).strip()


def 書き換える(d):
    """GAS の handleUpdateBlog。

    **送られてこなかった項目は触らない。**GAS も `data.title` があるときだけ
    書いている。全部を上書きにすると、管理アプリが一部だけ送ってきたときに
    残りが消える。
    """
    n = _探す(d)
    if not n:
        return {"status": "error", "message": "お知らせが見つかりませんでした"}

    if d.get("date"):
        n.posted_on = _日付(d.get("date")) or n.posted_on
    if d.get("title"):
        n.title = _文(d["title"]).strip()
    if d.get("category"):
        n.category = _文(d["category"]).strip()
    if d.get("icon"):
        n.icon = _文(d["icon"]).strip()
    if "body" in d:
        n.body = _文(d.get("body"))
    if d.get("status"):
        n.published = _公開か(d["status"])
    if d.get("noticeStatus"):
        n.notice_listed = _公開か(d["noticeStatus"])
    if "image" in d or "imageUrls" in d:
        n.image_url = _画像(d)
    if "publishAt" in d:
        n.publish_at = _日時(d.get("publishAt"))
    if "linkUrl" in d:
        n.link_url = _文(d.get("linkUrl")).strip()
    if "linkButtonText" in d:
        n.button_text = _文(d.get("linkButtonText")).strip()
    n.updated_at = timezone.now()
    n.save()
    return {"status": "ok"}


def 公開を変える(d):
    """GAS の handleUpdateRecordStatus のうち、お知らせのぶん。"""
    n = _探す(d)
    if not n:
        return {"status": "error", "message": "お知らせが見つかりませんでした"}
    n.published = _公開か(d.get("status"))
    n.updated_at = timezone.now()
    n.save(update_fields=["published", "updated_at", "changed_at"])
    return {"status": "ok"}


def 一覧掲載を変える(d):
    """GAS の handleUpdateNoticeVisibility。アプリの「お知らせ」に出すかどうか。"""
    n = _探す(d)
    if not n:
        return {"status": "error", "message": "お知らせが見つかりませんでした"}
    出す = _公開か(d.get("status"))
    n.notice_listed = 出す
    if 出す:
        n.notice_listed_at = timezone.now()
    else:
        n.notice_delisted_at = timezone.now()
    n.save()
    return {"status": "ok"}


def 一覧から外す(d):
    """GAS の handleDeleteNoticeListing。記事は残し、一覧から下ろすだけ。"""
    n = _探す(d)
    if not n:
        return {"status": "error", "message": "お知らせが見つかりませんでした"}
    n.notice_listed = False
    n.notice_delisted_at = timezone.now()
    n.save(update_fields=["notice_listed", "notice_delisted_at", "changed_at"])
    return {"status": "ok"}


def 消す(d):
    """GAS の handleDeleteRow のうち、お知らせのぶん。**印を付けるだけ。**"""
    n = _探す(d)
    if not n:
        return {"status": "error", "message": "お知らせが見つかりませんでした"}
    n.deleted = True
    n.deleted_at = timezone.now()
    n.delete_reason = _文(d.get("reason")).strip()
    n.save(update_fields=["deleted", "deleted_at", "delete_reason", "changed_at"])
    return {"status": "ok"}


def まとめて消す(d):
    """GAS の handleDeleteRows。行番号の並びを受ける。"""
    if not お知らせ宛てか(d):
        return {"status": "error", "message": "お知らせ以外の表は、まだサーバーにありません",
                "notImplemented": True}
    番号 = d.get("rowIdxs") or d.get("rows") or []
    if not isinstance(番号, list):
        return {"status": "error", "message": "行番号の並びが必要です"}
    数 = 0
    for r in 番号:
        結果 = 消す({"rowIdx": r, "sheet": d.get("sheet"), "reason": d.get("reason")})
        if 結果.get("status") == "ok":
            数 += 1
    return {"status": "ok", "deleted": 数}


def お知らせ宛てか(d):
    """`sheet` が指定されていて、お知らせ以外なら False。

    `deleteRow` / `updateRecordStatus` は**表をまたいで使い回されている**窓口。
    GAS 側で振り分けてから渡す作りだが、**ここでも確かめる。**
    渡す先を間違えたとき、黙って別の表のつもりの行番号でお知らせを消すと、
    気づくのが遅れる。
    """
    表 = _文(d.get("sheet")).strip().upper()
    return (not 表) or 表 == "BLOG"


def _探す(d):
    """**行番号で探す。**タイトルでは探さない（同じ題の記事があるため）。"""
    if not お知らせ宛てか(d):
        return None
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        return None
    if not 行:
        return None
    return News.objects.filter(sheet_row=行).first()
