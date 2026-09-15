from django.contrib.auth import views as auth_views
from django.urls import path
from django.views.generic import RedirectView

from . import views_calendar, views_category, views_login, views_menu, views_news, views_order, views_product, views_push

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
    path("menus/", views_menu.menu_list, name="menu_list"),
    path("menus/new/", views_menu.menu_create, name="menu_create"),
    path("menus/<int:row>/", views_menu.menu_edit, name="menu_edit"),
    path("menus/<int:row>/toggle/", views_menu.menu_toggle, name="menu_toggle"),
    path("menus/<int:row>/move/", views_menu.menu_move, name="menu_move"),
    path("menus/<int:row>/delete/", views_menu.menu_delete, name="menu_delete"),
    path("products/", views_product.product_list, name="product_list"),
    path("products/new/", views_product.product_create, name="product_create"),
    path("products/<int:row>/", views_product.product_edit, name="product_edit"),
    path("products/<int:row>/toggle/", views_product.product_toggle, name="product_toggle"),
    path("products/<int:row>/sold-out/", views_product.product_sold_out, name="product_sold_out"),
    path("products/<int:row>/delete/", views_product.product_delete, name="product_delete"),
    path("calendar/", views_calendar.calendar_list, name="calendar_list"),
    path("calendar/new/", views_calendar.calendar_create, name="calendar_create"),
    path("calendar/<int:row>/", views_calendar.calendar_edit, name="calendar_edit"),
    path("calendar/<int:row>/toggle/", views_calendar.calendar_toggle, name="calendar_toggle"),
    path("calendar/<int:row>/delete/", views_calendar.calendar_delete, name="calendar_delete"),
    path("orders/", views_order.order_list, name="order_list"),
    path("orders/new/", views_order.order_create, name="order_create"),
    path("orders/<str:order_id>/edit/", views_order.order_edit, name="order_edit"),
    path("orders/<str:order_id>/status/", views_order.order_status, name="order_status"),
    path("orders/<str:order_id>/delete/", views_order.order_delete, name="order_delete"),
]
