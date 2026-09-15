from django.urls import path, re_path

from . import views, views_basic, views_blog, views_classes, views_closed, views_images, views_layout, views_news, views_options

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
    path("hp/closed/", views_closed.hp_closed, name="hp_closed"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/images/", views_images.hp_images, name="hp_images"),
    path("hp/classes/", views_classes.hp_classes, name="hp_classes"),
    path("hp/classes/<int:i>/", views_classes.hp_class_edit, name="hp_class_edit"),
    path("hp/layout/", views_layout.hp_layout, name="hp_layout"),
    path("hp/layout/data/", views_layout.hp_layout_data, name="hp_layout_data"),
    path("hp/layout/api/<slug:what>/", views_layout.hp_layout_api, name="hp_layout_api"),
    path("hp/layout/new-image/", views_layout.hp_layout_image, name="hp_layout_image"),
    path("hp/layout/pv/<slug:page>/", views_layout.hp_layout_view, name="hp_layout_view"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
