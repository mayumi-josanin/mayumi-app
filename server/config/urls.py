from django.conf import settings
from django.contrib import admin
from django.urls import include, path

from apps.gasapi import media

urlpatterns = [
    path("admin/", admin.site.urls),
    # お客様のプロフィール写真。**WhiteNoise では配れない**ので専用に持つ
    # （起動時に一度しか走査しないため、上げた直後のものが 404 になる）。
    path("media/avatars/<str:名前>", media.写真, name="avatar"),
    # 管理画面で上げたお知らせの画像。お客様アプリが直接ここを見に来る。
    path("media/news/<str:名前>", media.お知らせ画像, name="news-image"),
    path("media/menus/<str:名前>", media.メニュー画像, name="menu-image"),
    path("media/products/<str:名前>", media.商品画像, name="product-image"),
    path("media/calendar/<str:名前>", media.カレンダー画像, name="calendar-image"),
    path("api/", include("apps.measurements.urls")),
    # GAS と同じ形の窓口（?action=...）。切り替えはアプリの GAS_URL を
    # ここに向けるだけで済む。**measurements の後ろに置くこと。**
    # 先に置くと health / measurements まで拾ってしまう。
    path("api", include("apps.gasapi.urls")),
]

# 院の管理画面。**MANAGE_ENABLED を立てた箱だけ**が持つ。
# Funnel に出ている web には無い（404）ので、外からは辿れない。
if settings.MANAGE_ENABLED:
    urlpatterns.append(path("manage/", include("apps.manage.urls")))
