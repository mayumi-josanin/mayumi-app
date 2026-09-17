"""お客様の画面（/manage/preview/...）。3つのお客様アプリを枠の中で確かめる。

apps/manage/urls.py から include されるので、逆引きは manage:preview:app のように2段になる。
古い住所 /manage/bijiris/preview/ からはここへ転送している（apps/bijiris/urls.py）。
"""

from django.urls import path, re_path

from . import views

app_name = "preview"

urlpatterns = [
    path("app/", views.app, name="app"),
    path("bijiris/", views.bijiris, name="bijiris"),
    path("reserve/", views.reserve, name="reserve"),
    # いま作っている方のファイルを配る口。<どれ> は app か bijiris（予約システムは配らない）
    re_path(r"^files/(?P<kind>[a-z]+)/(?P<path>.*)$", views.files, name="files"),
]
