from django.contrib.auth import views as auth_views
from django.urls import path
from django.views.generic import RedirectView

from . import views_login, views_news

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
]
