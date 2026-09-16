"""開発（/manage/dev/...）。並びは KEM の apps/devkanri/urls.py と同じ。

apps/manage/urls.py から include されるので、逆引きは manage:dev:project_list のように2段になる。
"""

from django.urls import path

from . import views

app_name = "dev"

urlpatterns = [
    # プロジェクト
    path("", views.project_list, name="project_list"),
    path("new/", views.project_create, name="project_create"),
    path("<int:pk>/", views.project_detail, name="project_detail"),
    path("<int:pk>/edit/", views.project_edit, name="project_edit"),
    path("<int:pk>/delete/", views.project_delete, name="project_delete"),
    # タスク
    path("<int:project_pk>/tasks/new/", views.task_create, name="task_create"),
    path("tasks/<int:pk>/", views.task_detail, name="task_detail"),
    path("tasks/<int:pk>/edit/", views.task_edit, name="task_edit"),
    path("tasks/<int:pk>/delete/", views.task_delete, name="task_delete"),
    path("tasks/<int:pk>/move/", views.task_move, name="task_move"),
    # 目安箱
    path("meyasubako/", views.meyasubako_list, name="meyasubako_list"),
    path("meyasubako/new/", views.meyasubako_create, name="meyasubako_create"),
    path("meyasubako/<int:pk>/", views.meyasubako_detail, name="meyasubako_detail"),
    path("meyasubako/<int:pk>/resolve/", views.meyasubako_resolve, name="meyasubako_resolve"),
]
