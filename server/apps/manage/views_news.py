"""お知らせ管理。GitHub Pages の管理アプリ「NEWS配信・管理」を、ここへ移したもの。

GAS を通らず、この Django が直接 News の表を読み書きする。
お客様アプリは GAS → サーバーの getNews で同じ表を読むので、ここで直せば届く
（GAS のキャッシュぶん、最長10分の遅れはある）。

**行番号（sheet_row）を鍵として使い続ける。**旧管理アプリが rowIdx で指しており、
並行して使っている間に食い違わないようにするため。

通知（OneSignal）はまだ送らない。旧管理アプリの投稿では GAS が送っている。
移すまでは「通知を出したい投稿は旧管理アプリから」。
"""

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import Category, News
from apps.gasapi.views import _画像

from . import images
from .forms import NewsForm


def _種別(category_kinds: dict, name: str) -> str:
    kind = category_kinds.get((name or "").strip())
    if kind:
        return kind
    return "お知らせ" if name == "お知らせ" else "ブログ"


@login_required
def news_list(request):
    kinds = {c.name: (c.kind or "ブログ") for c in Category.objects.all()}
    filter_kind = request.GET.get("kind", "")
    filter_category = request.GET.get("category", "")

    rows = News.objects.filter(deleted=False).order_by("-posted_on", "-sheet_row")
    if filter_category:
        rows = rows.filter(category=filter_category)

    published, drafts = [], []
    for n in rows:
        kind = _種別(kinds, n.category)
        if filter_kind and kind != filter_kind:
            continue
        item = {
            "n": n,
            "kind": kind,
            "image": (_画像(n.image_url) or [""])[0],
        }
        (published if n.published else drafts).append(item)

    return render(
        request,
        "manage/news_list.html",
        {
            "published": published,
            "drafts": drafts,
            "categories": Category.objects.order_by("sheet_row"),
            "filter_kind": filter_kind,
            "filter_category": filter_category,
        },
    )


def _保存(request, form: NewsForm, n: News | None):
    """投稿と修正の共通部分。画像は「残すもの」＋「新しく上げたもの」。"""
    obj = form.save(commit=False)
    status = request.POST.get("status") or "公開"
    obj.published = status != "非公開"

    kept = [u for u in request.POST.getlist("keep_image") if u.strip()]
    added = [images.保存する(f) for f in form.cleaned_data.get("images") or []]
    # views.py の _画像 は改行区切りも JSON の配列も解けるので、改行でつなぐ（admin_news と同じ）
    obj.image_url = "\n".join(kept + added)
    obj.updated_at = timezone.now()

    if n is None:
        with transaction.atomic():
            # 新しい行番号は、いまの最大＋1（シートの appendRow と同じ）。
            最大 = (
                News.objects.select_for_update().order_by("-sheet_row")
                .values_list("sheet_row", flat=True).first() or 1
            )
            obj.sheet_row = 最大 + 1
            obj.save()
    else:
        obj.save()
    return obj


@login_required
def news_create(request):
    if request.method == "POST":
        form = NewsForm(request.POST, request.FILES)
        if form.is_valid():
            obj = _保存(request, form, None)
            messages.success(
                request,
                "下書きとして保存しました。" if not obj.published else "お知らせを投稿しました。",
            )
            return redirect("manage:news_list")
    else:
        form = NewsForm(initial={"posted_on": timezone.localdate(), "icon": "📢", "notice_listed": True})
    return render(
        request,
        "manage/news_form.html",
        {"form": form, "news": None, "existing_images": []},
    )


@login_required
def news_edit(request, row: int):
    n = get_object_or_404(News, sheet_row=row, deleted=False)
    if request.method == "POST":
        form = NewsForm(request.POST, request.FILES, instance=n)
        if form.is_valid():
            obj = _保存(request, form, n)
            messages.success(
                request,
                "下書きとして保存しました。" if not obj.published else "お知らせを更新しました。",
            )
            return redirect("manage:news_list")
    else:
        form = NewsForm(instance=n)
    return render(
        request,
        "manage/news_form.html",
        {"form": form, "news": n, "existing_images": _画像(n.image_url)},
    )


@login_required
@require_POST
def news_toggle(request, row: int):
    """公開 ⇄ 非公開。"""
    n = get_object_or_404(News, sheet_row=row, deleted=False)
    n.published = not n.published
    n.updated_at = timezone.now()
    n.save(update_fields=["published", "updated_at", "changed_at"])
    messages.success(request, f"「{n.title}」を{'公開' if n.published else '非公開'}にしました。")
    return redirect("manage:news_list")


@login_required
@require_POST
def news_delete(request, row: int):
    """消さずに印を付ける（GAS と同じ論理削除）。過去の掲載についての問い合わせは実際に来る。"""
    n = get_object_or_404(News, sheet_row=row, deleted=False)
    n.deleted = True
    n.deleted_at = timezone.now()
    n.delete_reason = f"管理画面から（{request.user.username}）"
    n.updated_at = timezone.now()
    n.save(update_fields=["deleted", "deleted_at", "delete_reason", "updated_at", "changed_at"])
    messages.success(request, f"「{n.title}」を削除しました。")
    return redirect("manage:news_list")
