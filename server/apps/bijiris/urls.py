"""ビジリス管理（/manage/bijiris/...）。院長が決めた6画面（集計／アンケート管理／回答管理／顧客管理／特典／回数券分析）。

apps/manage/urls.py から include されるので、逆引きは manage:bijiris:dashboard のように2段になる。
**中身はまだ「準備中」**（段取り A）。段取り B で1画面ずつ置き換える。URL 名は変えない（左メニューが当てている）。
"""

from django.urls import path

from . import views, views_ticket

app_name = "bijiris"

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("surveys/", views.survey_list, name="survey_list"),
    path("responses/", views.response_list, name="response_list"),
    path("customers/", views.customer_list, name="customer_list"),
    path("rewards/", views.reward_list, name="reward_list"),
    path("tickets/", views_ticket.ticket_list, name="ticket_list"),
    path("tickets/analyze/", views_ticket.ticket_analyze, name="ticket_analyze"),
    path("tickets/prompt/", views_ticket.ticket_prompt_save, name="ticket_prompt"),
    path("tickets/auto/", views_ticket.ticket_auto, name="ticket_auto"),
]
