"""経費管理（/manage/keihi/...）。

apps/manage/urls.py から include されるので、逆引きは manage:keihi:book のように2段になる。
"""

from django.urls import path

from . import views

app_name = "keihi"

urlpatterns = [
    path("", views.book, name="book"),
    path("csv/", views.csv_export, name="csv_export"),
    path("csv/import/", views.csv_import, name="csv_import"),
]
