from django.urls import path, re_path

from . import views, views_classes

urlpatterns = [
    path("hp/git/", views.hp_git, name="hp_git"),
    path("hp/publish/", views.hp_publish, name="hp_publish"),
    path("hp/classes/", views_classes.hp_classes, name="hp_classes"),
    path("hp/classes/<int:i>/", views_classes.hp_class_edit, name="hp_class_edit"),
    path("hp/preview/", views.hp_preview, name="hp_preview"),
    re_path(r"^hp/preview/files/(?P<path>.*)$", views.hp_preview_file, name="hp_preview_file"),
    re_path(r"^hp/site/(?P<path>.*)$", views.hp_site_file, name="hp_site_file"),
]
