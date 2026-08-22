from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    path("api/", include("apps.measurements.urls")),
    # GAS と同じ形の窓口（?action=...）。切り替えはアプリの GAS_URL を
    # ここに向けるだけで済む。**measurements の後ろに置くこと。**
    # 先に置くと health / measurements まで拾ってしまう。
    path("api", include("apps.gasapi.urls")),
]
