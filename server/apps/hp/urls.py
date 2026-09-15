from django.urls import path, re_path

from . import views, views_basic

urlpatterns = [
    path("hp/basic/", views_basic.hp_basic, name="hp_basic"),
    path("hp/basic/save/", views_basic.hp_basic_save, name="hp_basic_save"),
    path("hp/basic/price-preview/", views_basic.hp_basic_price_preview, name="hp_basic_price_preview"),
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
