"""ビジリス管理（/manage/bijiris/...）。院長が決めた6画面（集計／アンケート管理／回答管理／顧客管理／特典／回数券分析）。

apps/manage/urls.py から include されるので、逆引きは manage:bijiris:dashboard のように2段になる。
段取り B で1画面ずつ置き換える。URL 名は変えない（左メニューが当てている）。
「お客様の画面」は apps/preview へ引っ越したので、ここには転送だけが残っている。
"""

from django.shortcuts import redirect
from django.urls import path, re_path
from django.views.generic import RedirectView

from . import views_customers, views_dashboard, views_responses, views_rewards, views_surveys, views_ticket

app_name = "bijiris"


def _古いファイルの口(request, path=""):
    """古いファイルの配り口。新しい配り口（/manage/preview/files/bijiris/...）へ送る。"""
    return redirect(f"/manage/preview/files/bijiris/{path}")


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
    # 顧客管理（views_customers）。動作の path は詳細より前に置く（<str:name> に呑まれないように）
    path("customers/", views_customers.customer_list, name="customer_list"),
    path("customers/add/", views_customers.customer_add, name="customer_add"),
    path("customers/<str:name>/save/", views_customers.customer_save, name="customer_save"),
    path("customers/<str:name>/delete/", views_customers.customer_delete, name="customer_delete"),
    path("customers/<str:name>/passcode/", views_customers.customer_passcode_setup, name="customer_passcode_setup"),
    path("customers/<str:name>/memo/", views_customers.customer_memo, name="customer_memo"),
    path("customers/<str:name>/stamp/", views_customers.customer_stamp_adjust, name="customer_stamp_adjust"),
    path("customers/<str:name>/redemption/", views_customers.customer_redemption, name="customer_redemption"),
    path("customers/<str:name>/link/", views_customers.customer_link_member, name="customer_link_member"),
    path("customers/<str:name>/", views_customers.customer_detail, name="customer_detail"),
    # 特典（views_rewards）
    path("rewards/", views_rewards.reward_list, name="reward_list"),
    path("rewards/milestone/", views_rewards.reward_milestone_save, name="reward_milestone_save"),
    path("rewards/gacha/", views_rewards.reward_gacha_save, name="reward_gacha_save"),
    path("rewards/campaign/", views_rewards.reward_campaign_save, name="reward_campaign_save"),
    # 回数券分析（views_ticket）
    path("tickets/", views_ticket.ticket_list, name="ticket_list"),
    path("tickets/analyze/", views_ticket.ticket_analyze, name="ticket_analyze"),
    path("tickets/prompt/", views_ticket.ticket_prompt_save, name="ticket_prompt"),
    path("tickets/auto/", views_ticket.ticket_auto, name="ticket_auto"),
    # お客様の画面は apps/preview へ引っ越した（2026-09-18。3つのお客様アプリを1か所にまとめたため）。
    # 古い住所を覚えている人・ブックマークのために、新しい住所へ転送するだけ残す
    path("preview/", RedirectView.as_view(pattern_name="manage:preview:bijiris", permanent=False), name="preview"),
    re_path(r"^preview/files/(?P<path>.*)$", _古いファイルの口, name="preview_file"),
]
