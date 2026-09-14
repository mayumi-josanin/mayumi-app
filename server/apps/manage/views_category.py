"""カテゴリ管理。お知らせ・商品・メニューのカテゴリの表。

読み書きは gasapi/admin_category.py の関数をそのまま使う。
旧管理アプリ（GAS 経由）と同じ決まり（名前が鍵・名前を変えても記事は変えない・
同じ名前は足せない）を、2か所に書かないため。
"""

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Count
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from apps.content.models import Category, News
from apps.gasapi import admin_category

KINDS = admin_category.種別


@login_required
def category_list(request):
    使用数 = dict(
        News.objects.filter(deleted=False).values_list("category").annotate(n=Count("id"))
    )
    rows = [
        {"c": c, "used": 使用数.get(c.name, 0)}
        for c in Category.objects.order_by("sheet_row")
    ]
    return render(request, "manage/category_list.html", {"rows": rows, "kinds": KINDS})


@login_required
@require_POST
def category_add(request):
    答 = admin_category.足す(
        {"name": request.POST.get("name", ""), "categoryType": request.POST.get("kind", "")}
    )
    _知らせる(request, 答)
    return redirect("manage:category_list")


@login_required
@require_POST
def category_update(request):
    答 = admin_category.書き換える(
        {
            "oldName": request.POST.get("old_name", ""),
            "newName": request.POST.get("name", ""),
            "categoryType": request.POST.get("kind", ""),
        }
    )
    _知らせる(request, 答)
    return redirect("manage:category_list")


@login_required
@require_POST
def category_delete(request):
    答 = admin_category.消す({"name": request.POST.get("name", "")})
    _知らせる(request, 答)
    return redirect("manage:category_list")


def _知らせる(request, 答):
    if 答.get("status") == "ok":
        messages.success(request, 答.get("message") or "保存しました。")
    else:
        messages.error(request, 答.get("message") or "保存できませんでした。")
