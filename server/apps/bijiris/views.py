"""ビジリス管理の画面。**いまは全部「準備中」のカード**（段取り A）。

段取り B で1画面ずつ中身に置き換える。置き換えるときの決まり:
  - 全部 **まゆみだけ**（owner_required）。スタッフは 403
  - 読むのはいつでもよい。**書くのは gate.ビジリスはサーバーが正() のときだけ**。立っていなければ gate.断る文() を出して断る
  - お名前で会員を探さない（CLAUDE.md）
"""

from django.shortcuts import render

from apps.manage.permissions import owner_required
from apps.records.models import TicketAnalysis

from . import gate
from .models import CustomerProfile, Response, ResponsePhoto, Survey

画面 = {
    "dashboard": "集計",
    "survey_list": "アンケート管理",
    "response_list": "回答管理",
    "customer_list": "顧客管理",
    "reward_list": "特典",
    "ticket_list": "回数券分析",
}


def _準備中(request, 名):
    件数 = [
        ("アンケート", Survey.objects.count()),
        ("回答", Response.objects.count()),
        ("回答の写真", ResponsePhoto.objects.count()),
        ("顧客", CustomerProfile.objects.count()),
        ("回数券分析", TicketAnalysis.objects.count()),
    ]
    return render(request, "bijiris/placeholder.html", {
        "title": 画面[名],
        "server_is_source": gate.ビジリスはサーバーが正(),
        "refuse_text": gate.断る文(),
        "counts": 件数,
    })


@owner_required
def dashboard(request):
    return _準備中(request, "dashboard")


@owner_required
def survey_list(request):
    return _準備中(request, "survey_list")


@owner_required
def response_list(request):
    return _準備中(request, "response_list")


@owner_required
def customer_list(request):
    return _準備中(request, "customer_list")


@owner_required
def reward_list(request):
    return _準備中(request, "reward_list")


@owner_required
def ticket_list(request):
    return _準備中(request, "ticket_list")
