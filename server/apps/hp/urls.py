from django.urls import path, re_path

from . import views, views_news

urlpatterns = [
    path("hp/news/", views_news.hp_news, name="hp_news"),
    path("hp/news/app/", views_news.hp_news_app, name="hp_news_app"),
    path("hp/news/<int:index>/", views_news.hp_news_edit, name="hp_news_edit"),
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
