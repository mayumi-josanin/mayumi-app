"""役割を選んで、パスコードで入る。予約システムと同じ作り。

**お名前を打たせない。**院で使う端末は決まっていて、入るのは
まゆみさんかスタッフのどちらかしかいない。

パスコードは、その役割のアカウントのパスワードそのもの。別に持たない。
同じ役割に何人か登録してあるときは、合った方として入る。
"""

import unicodedata

from django.contrib.auth import login
from django.contrib.auth.models import User
from django.shortcuts import redirect, render
from django.urls import reverse

from .permissions import OWNER, ROLES, STAFF


def _candidates(role: str):
    people = User.objects.filter(is_active=True)
    if role == OWNER:
        return people.filter(groups__name=OWNER) | people.filter(is_superuser=True)
    return people.filter(groups__name=role)


def role_login(request):
    role = request.POST.get("role") or request.GET.get("role") or ""
    if role not in ROLES:
        role = ""
    error = ""
    if request.method == "POST":
        # 日本語の入力では「０７１４」と全角になる。ここで弾くと
        # 「合っているはずなのに入れない」になる。
        passcode = unicodedata.normalize("NFKC", request.POST.get("passcode") or "").strip()
        if not role:
            error = "まゆみかスタッフかを選んでください。"
        elif not passcode:
            error = "パスコードを入れてください。"
        else:
            for user in _candidates(role).distinct():
                if user.check_password(passcode):
                    login(request, user)
                    next_url = request.POST.get("next") or ""
                    # 外の住所へ飛ばされないよう、この画面の中だけを許す。
                    if not next_url.startswith("/manage/"):
                        next_url = reverse("manage:news_list")
                    return redirect(next_url)
            error = "パスコードが違うようです。もう一度お試しください。"
    return render(
        request,
        "manage/login.html",
        {
            "role": role,
            "roles": ROLES,
            "owner": OWNER,
            "staff": STAFF,
            "error": error,
            "next": request.POST.get("next") or request.GET.get("next") or "",
        },
    )
