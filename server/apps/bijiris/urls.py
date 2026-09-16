"""ビジリス管理（/manage/bijiris/...）。院長が決めた6画面（集計／アンケート管理／回答管理／顧客管理／特典／回数券分析）。

apps/manage/urls.py から include されるので、逆引きは manage:bijiris:dashboard のように2段になる。
**中身はまだ「準備中」**（段取り A）。段取り B で1画面ずつ置き換える。URL 名は変えない（左メニューが当てている）。
"""

from django.urls import path

from . import views, views_customers, views_rewards, views_ticket

app_name = "bijiris"

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("surveys/", views.survey_list, name="survey_list"),
    path("responses/", views.response_list, name="response_list"),
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
]
