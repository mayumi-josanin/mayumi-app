from django.contrib.auth import views as auth_views
from django.urls import path
from django.views.generic import RedirectView

from . import views_category, views_login, views_news, views_push

app_name = "manage"

urlpatterns = [
    path("login/", views_login.role_login, name="login"),
    # ログアウトはボタン（POST）。リンクだと先読みで勝手に落ちることがある。
    path("logout/", auth_views.LogoutView.as_view(next_page="/manage/login/"), name="logout"),
    path("", RedirectView.as_view(pattern_name="manage:news_list", permanent=False), name="home"),
    path("news/", views_news.news_list, name="news_list"),
    path("news/new/", views_news.news_create, name="news_create"),
    path("news/<int:row>/", views_news.news_edit, name="news_edit"),
    path("news/<int:row>/toggle/", views_news.news_toggle, name="news_toggle"),
    path("news/<int:row>/delete/", views_news.news_delete, name="news_delete"),
    path("categories/", views_category.category_list, name="category_list"),
    path("categories/add/", views_category.category_add, name="category_add"),
    path("categories/update/", views_category.category_update, name="category_update"),
    path("categories/delete/", views_category.category_delete, name="category_delete"),
    path("push/", views_push.push_list, name="push_list"),
    path("push/send/", views_push.push_send, name="push_send"),
    path("push/<int:row>/delete/", views_push.push_delete, name="push_delete"),
]
