from django.urls import path, re_path

from . import views, views_basic, views_blog, views_images, views_news, views_options

urlpatterns = [
    path("hp/options/", views_options.hp_options, name="hp_options"),
    path("hp/blog/", views_blog.hp_blog, name="hp_blog"),
    path("hp/blog/<slug:slug>/", views_blog.hp_blog_edit, name="hp_blog_edit"),
    path("hp/basic/", views_basic.hp_basic, name="hp_basic"),
    path("hp/basic/save/", views_basic.hp_basic_save, name="hp_basic_save"),
    path("hp/basic/price-preview/", views_basic.hp_basic_price_preview, name="hp_basic_price_preview"),
    path("hp/news/", views_news.hp_news, name="hp_news"),
    path("hp/news/app/", views_news.hp_news_app, name="hp_news_app"),
    path("hp/news/<int:index>/", views_news.hp_news_edit, name="hp_news_edit"),
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/images/", views_images.hp_images, name="hp_images"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
