from django.contrib import admin
from django.urls import include, path

from apps.gasapi import media

urlpatterns = [
    path("admin/", admin.site.urls),
    # お客様のプロフィール写真。**WhiteNoise では配れない**ので専用に持つ
    # （起動時に一度しか走査しないため、上げた直後のものが 404 になる）。
    path("media/avatars/<str:名前>", media.写真, name="avatar"),
    path("api/", include("apps.measurements.urls")),
    # GAS と同じ形の窓口（?action=...）。切り替えはアプリの GAS_URL を
    # ここに向けるだけで済む。**measurements の後ろに置くこと。**
    # 先に置くと health / measurements まで拾ってしまう。
    path("api", include("apps.gasapi.urls")),
]
