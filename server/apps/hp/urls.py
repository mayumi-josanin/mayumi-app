from django.urls import path, re_path

from . import views, views_layout

urlpatterns = [
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/layout/", views_layout.hp_layout, name="hp_layout"),
    path("hp/layout/data/", views_layout.hp_layout_data, name="hp_layout_data"),
    path("hp/layout/api/<slug:what>/", views_layout.hp_layout_api, name="hp_layout_api"),
    path("hp/layout/new-image/", views_layout.hp_layout_image, name="hp_layout_image"),
    path("hp/layout/pv/<slug:page>/", views_layout.hp_layout_view, name="hp_layout_view"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
