"""NEWS配信・管理。**中身は旧管理アプリ（#page-blogs NEWS配信・管理）と同じ。**

読み書きは gasapi/admin_news.py（GAS の転送先と同じ関数群）。ここで書き込み経路を増やさない。
お客様アプリは GAS → サーバーの getNews で同じ表を読むので、ここで直せば届く
（GAS のキャッシュぶん、最長10分の遅れはある）。

**行番号（sheet_row）を鍵として使い続ける。**旧管理アプリが rowIdx で指しており、
並行して使っている間に食い違わないようにするため。

通知（OneSignal）は「この保存でPush通知を送信する」にチェックがあるときだけ、
このサーバーから送る（push.py）。条件は GAS の _お知らせの通知を送る_ と同じ
（公開保存で、公開開始日時が来ているときだけ）。
"""

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Category, News
from apps.gasapi import admin_news
from apps.gasapi.admin_news import _日時
from apps.gasapi.views import _画像

from . import images, push
from .permissions import owner_required

# 旧管理アプリの「カテゴリ」select に最初から入っている候補。
# カテゴリ管理に「お知らせ」「ブログ」の分類が1つも無いときだけ、これを出す。
既定のカテゴリ = ["お知らせ", "ブログ", "休診情報", "イベント", "商品情報"]


def _種別(kinds: dict, name: str) -> str:
    """カテゴリ名から「お知らせ」か「ブログ」か。旧アプリの loadAdminBlogs と同じ決め方:
    カテゴリ管理に種別があればそれ、無ければ名前が「お知らせ」のときだけお知らせ。"""
    kind = (kinds.get((name or "").strip()) or "").strip()
    if kind:
        return kind
    return "お知らせ" if name == "お知らせ" else "ブログ"


def _NEWSのカテゴリ():
    """カテゴリ管理のうち NEWS に使うもの（お知らせ／ブログ）。種別が空ならブログ扱い
    （旧アプリの updateAllCategorySelectors と同じ）。[(名前, 種別), ...] を表の順で返す。"""
    出 = []
    for c in Category.objects.order_by("sheet_row"):
        kind = (c.kind or "").strip() or "ブログ"
        if kind in ("お知らせ", "ブログ"):
            出.append((c.name, kind))
    return 出


def _一件(b: dict, kinds: dict) -> dict:
    """一覧の1行。アイコン欄は、アイコンがURLならそれ、無ければ1枚目の画像、それも無ければ絵文字
    （旧アプリの buildAdminBlogIconHtml と同じ）。"""
    icon = (b.get("icon") or "").strip()
    if icon.lower().startswith(("http://", "https://", "data:image/")):
        image = icon
    else:
        image = (b.get("imageUrls") or [""])[0]
    return {"b": b, "kind": _種別(kinds, b.get("category")), "image": image,
            "icon": icon if not image else ""}


@owner_required
def news_list(request):
    kinds = {c.name: (c.kind or "") for c in Category.objects.all()}
    filter_kind = request.GET.get("kind", "")
    filter_category = request.GET.get("category", "")

    # 新しい順（日付の新しい順、同じ日付なら行番号の大きいほう）。旧アプリの loadAdminBlogs と同じ。
    # 一覧() はシートの順（行番号順）で返してくるので、ここで並べ替える。
    blogs = sorted(admin_news.一覧()["blogs"], key=lambda b: (b["date"] or "", b["rowIdx"]), reverse=True)

    drafts, published = [], []
    for b in blogs:
        item = _一件(b, kinds)
        # 絞り込みは下書き・投稿済みの両方に効く（旧アプリと同じ）
        if filter_category and b["category"] != filter_category:
            continue
        if filter_kind and item["kind"] != filter_kind:
            continue
        (drafts if b["status"] == "非公開" else published).append(item)

    return render(request, "manage/news_list.html", {
        "drafts": drafts, "published": published,
        "categories": [name for name, _ in _NEWSのカテゴリ()],
        "filter_kind": filter_kind, "filter_category": filter_category,
    })


def _入力(request, n: News | None) -> dict:
    """フォームの値を、GAS の addBlog / updateBlog が受ける形に。画像は「残すもの」＋「新しく上げたもの」。"""
    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f, "news") for f in request.FILES.getlist("images")]
    urls = kept + added
    return {
        "rowIdx": n.sheet_row if n else 0,
        "date": request.POST.get("date", "").strip(),
        "title": request.POST.get("title", "").strip(),
        "category": request.POST.get("category", "").strip(),
        "icon": request.POST.get("icon", "").strip(),
        "image": urls[0] if urls else "",
        "imageUrls": urls,
        "body": request.POST.get("body", ""),
        "status": request.POST.get("status") or "公開",
        "publishAt": request.POST.get("publishAt", "").strip(),
        "linkUrl": request.POST.get("linkUrl", "").strip(),
        "linkButtonText": request.POST.get("linkButtonText", "").strip(),
    }


def _通知(request, 答: dict, d: dict, 題: str):
    """GAS の _お知らせの通知を送る_ と同じ条件: チェックあり・公開・公開開始日時が来ている。"""
    if not request.POST.get("send_push"):
        return
    状態 = 答.get("effectiveStatus") or d.get("status") or "公開"
    if 状態 == "非公開":
        return
    公開日時 = 答.get("effectivePublishAt") if "effectivePublishAt" in 答 else d.get("publishAt")
    t = _日時(公開日時) if 公開日時 else None
    if t and t > timezone.now():
        return
    if push.全員へ送る(題, "NEWSが更新されました", page="news"):
        messages.info(request, "お客様のアプリへ通知を送りました。")
    else:
        messages.error(request, "通知を送れませんでした（投稿は保存されています）。")


def _お知らせ一覧へ反映(request, 答: dict, d: dict, row: int):
    """「この保存をお知らせ一覧に反映する」。チェックがあり、公開保存のときだけ、
    お客様アプリの「お知らせ」一覧に載せ直す（掲載日時を今にする）。未チェックなら通常保存のみ。"""
    if not request.POST.get("notice_refresh"):
        return
    if (答.get("effectiveStatus") or d.get("status") or "公開") == "非公開":
        return
    admin_news.一覧掲載を変える({"rowIdx": row, "sheet": "BLOG", "status": "公開"})


def _値(n: News) -> dict:
    """修正画面の初期値。日時は「更新日時」（無ければ投稿日）。旧アプリの editBlog と同じ。"""
    if n.updated_at:
        date = timezone.localtime(n.updated_at).strftime("%Y-%m-%dT%H:%M")
    elif n.posted_on:
        date = n.posted_on.strftime("%Y-%m-%dT00:00")
    else:
        date = ""
    return {
        "date": date, "title": n.title, "category": n.category, "icon": n.icon or "📢", "body": n.body,
        "publishAt": timezone.localtime(n.publish_at).strftime("%Y-%m-%dT%H:%M") if n.publish_at else "",
        "linkUrl": n.link_url, "linkButtonText": n.button_text,
    }


def _画面(request, n, values, existing, notice_refresh: bool):
    """notice_refresh: 「お知らせ一覧に反映する」の初期値。新規はオン、修正はオフ（旧アプリと同じ）。"""
    候補 = _NEWSのカテゴリ()
    groups = [
        ("お知らせ分類", [name for name, kind in 候補 if kind == "お知らせ"]),
        ("ブログ分類", [name for name, kind in 候補 if kind == "ブログ"]),
    ]
    groups = [g for g in groups if g[1]]
    current = (values.get("category") or "").strip()
    names = [name for name, _ in 候補]
    extra = []
    if not names:
        extra = list(既定のカテゴリ)
    if current and current not in names and current not in extra:
        extra.append(current)  # いまの値が候補に無くても消さない
    return render(request, "manage/news_form.html", {
        "news": n, "values": values, "existing_images": existing,
        "category_groups": groups, "category_extra": extra, "notice_refresh": notice_refresh,
    })


@owner_required
def news_create(request):
    if request.method == "POST":
        d = _入力(request, None)
        答 = admin_news.足す(d)
        if 答.get("status") == "ok":
            messages.success(request, "お知らせを投稿しました！ 🎉")
            _お知らせ一覧へ反映(request, 答, d, 答["rowIdx"])
            _通知(request, 答, d, "📝 " + (d["title"] or "新しい投稿"))
            return redirect("manage:news_list")
        messages.error(request, 答.get("message") or "投稿に失敗しました")
        return _画面(request, None, request.POST, [], bool(request.POST.get("notice_refresh")))
    now = timezone.localtime(timezone.now()).strftime("%Y-%m-%dT%H:%M")
    return _画面(request, None, {"date": now, "icon": "📢"}, [], True)


@owner_required
def news_edit(request, row: int):
    n = get_object_or_404(News, sheet_row=row, deleted=False)
    if request.method == "POST":
        d = _入力(request, n)
        if not d["title"]:
            messages.error(request, "タイトルが必要です")
            return _画面(request, n, request.POST, _画像(n.image_url), bool(request.POST.get("notice_refresh")))
        答 = admin_news.書き換える(d)
        if 答.get("status") == "ok":
            messages.success(request, "ブログを更新しました！")
            _お知らせ一覧へ反映(request, 答, d, row)
            _通知(request, 答, d, "📝 " + (d["title"] or "ブログ更新"))
            return redirect("manage:news_list")
        messages.error(request, 答.get("message") or "投稿に失敗しました")
        return _画面(request, n, request.POST, _画像(n.image_url), bool(request.POST.get("notice_refresh")))
    return _画面(request, n, _値(n), _画像(n.image_url), False)


@owner_required
@require_POST
def news_status(request, row: int):
    """一覧の「公開状況」select（公開／非公開）。旧アプリの updateBlogStatus と同じ。"""
    get_object_or_404(News, sheet_row=row, deleted=False)
    答 = admin_news.公開を変える({"rowIdx": row, "sheet": "BLOG", "status": request.POST.get("status", "公開")})
    if 答.get("status") == "ok":
        messages.success(request, "公開設定を変更しました")
    else:
        messages.error(request, 答.get("message") or "設定を変えられませんでした。")
    return redirect(request.POST.get("next") or "manage:news_list")


@owner_required
@require_POST
def news_delete(request, row: int):
    """消さずに印を付ける（GAS と同じ論理削除）。過去の掲載についての問い合わせは実際に来る。"""
    get_object_or_404(News, sheet_row=row, deleted=False)
    答 = admin_news.消す({"rowIdx": row, "sheet": "BLOG", "reason": f"管理画面から（{request.user.username}）"})
    if 答.get("status") == "ok":
        messages.success(request, "ブログを削除しました")
    else:
        messages.error(request, "削除に失敗しました")
    return redirect(request.POST.get("next") or "manage:news_list")


@owner_required
@require_POST
def news_bulk_delete(request):
    rows = [int(x) for x in request.POST.getlist("rows") if x.isdigit()]
    if not rows:
        messages.error(request, "削除する項目を選択してください。")
        return redirect(request.POST.get("next") or "manage:news_list")
    答 = admin_news.まとめて消す({"rowIdxs": rows, "sheet": "BLOG", "reason": f"管理画面から一括（{request.user.username}）"})
    messages.success(request, f"{答.get('deleted', 0)}件のブログ記事を削除しました")
    return redirect(request.POST.get("next") or "manage:news_list")
