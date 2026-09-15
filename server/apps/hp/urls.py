from django.urls import path, re_path

from . import views, views_blog

urlpatterns = [
    path("hp/blog/", views_blog.hp_blog, name="hp_blog"),
    path("hp/blog/<slug:slug>/", views_blog.hp_blog_edit, name="hp_blog_edit"),
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
