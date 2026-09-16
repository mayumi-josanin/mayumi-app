"""ビジリス管理（/manage/bijiris/...）。院長が決めた6画面（集計／アンケート管理／回答管理／顧客管理／特典／回数券分析）。

apps/manage/urls.py から include されるので、逆引きは manage:bijiris:dashboard のように2段になる。
段取り B で1画面ずつ置き換える。URL 名は変えない（左メニューが当てている）。
"""

from django.urls import path

from . import views, views_dashboard, views_responses, views_surveys

app_name = "bijiris"

urlpatterns = [
    # 集計（views_dashboard）
    path("", views_dashboard.dashboard, name="dashboard"),
    # アンケート管理（views_surveys）
    path("surveys/", views_surveys.survey_list, name="survey_list"),
    path("surveys/new/", views_surveys.survey_create, name="survey_create"),
    path("surveys/<str:survey_id>/edit/", views_surveys.survey_edit, name="survey_edit"),
    path("surveys/<str:survey_id>/move/", views_surveys.survey_move, name="survey_move"),
    path("surveys/<str:survey_id>/duplicate/", views_surveys.survey_duplicate, name="survey_duplicate"),
    path("surveys/<str:survey_id>/archive/", views_surveys.survey_archive, name="survey_archive"),
    path("surveys/<str:survey_id>/delete/", views_surveys.survey_delete, name="survey_delete"),
    # 回答管理（views_responses）
    path("responses/", views_responses.response_list, name="response_list"),
    path("responses/export.csv", views_responses.response_export, name="response_export"),
    path("responses/print/", views_responses.response_print_list, name="response_print_list"),
    path("responses/<str:response_id>/", views_responses.response_detail, name="response_detail"),
    path("responses/<str:response_id>/print/", views_responses.response_print, name="response_print"),
    path("responses/<str:response_id>/save/", views_responses.response_save, name="response_save"),
    path("responses/<str:response_id>/trash/", views_responses.response_trash, name="response_trash"),
    path("responses/<str:response_id>/purge/", views_responses.response_purge, name="response_purge"),
    path("customers/", views.customer_list, name="customer_list"),
    path("rewards/", views.reward_list, name="reward_list"),
    path("tickets/", views.ticket_list, name="ticket_list"),
]
