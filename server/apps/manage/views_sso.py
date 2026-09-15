"""予約管理（mayumi-reserve）とのログインの受け渡し。

    /manage/go/reserve/?to=/manage/customers/   … 予約管理へ（札を付けて）
    /manage/sso/?t=<札>&next=/manage/orders/     … 予約管理から来た札で入る
"""

from urllib.parse import quote

from django.conf import settings
from django.contrib import messages
from django.contrib.auth import login
from django.contrib.auth.models import User
from django.core.cache import cache
from django.http import HttpResponseRedirect
from django.shortcuts import redirect

from . import sso
from .permissions import OWNER, ROLES, STAFF, is_owner, roles_of


def _中だけ(path: str, 既定: str) -> str:
    """相手の管理画面の中だけへ。外の住所や別の場所には飛ばさない。"""
    p = (path or "").strip()
    return p if p.startswith("/manage/") and not p.startswith("//") else 既定


def go_reserve(request):
    if not request.user.is_authenticated:
        return redirect(f"/manage/login/?next={quote(request.get_full_path())}")
    to = _中だけ(request.GET.get("to"), "/manage/")
    base = settings.RESERVE_MANAGE_URL.rstrip("/")
    if not settings.MANAGE_SSO_SECRET:
        # 合鍵が無ければ、ただ相手へ（相手のログイン画面に着く）
        return HttpResponseRedirect(base + to)
    role = OWNER if is_owner(request.user) else (STAFF if STAFF in roles_of(request.user) else "")
    if not role:
        return HttpResponseRedirect(base + to)
    札 = sso.札を作る(role, settings.MANAGE_SSO_SECRET)
    return HttpResponseRedirect(f"{base}/manage/sso/?t={quote(札)}&next={quote(to)}")


def sso_login(request):
    """予約管理から渡された札で入る。"""
    next_url = _中だけ(request.GET.get("next"), "/manage/")
    中 = sso.札を確かめる(request.GET.get("t", ""), settings.MANAGE_SSO_SECRET)
    if not 中 or 中["role"] not in ROLES:
        messages.error(request, "予約管理からの受け渡しが確かめられませんでした。ログインしてください。")
        return redirect(f"/manage/login/?next={quote(next_url)}")
    if 中["role"] != OWNER:
        # **スタッフはアプリ管理を見ない**（院長の決定 2026-09-15）
        messages.error(request, "アプリ管理はまゆみだけが使えます。")
        return redirect(f"/manage/login/?next={quote(next_url)}")
    # 同じ札を二度使わせない
    if not cache.add("manage-sso-nonce:" + 中["nonce"], 1, sso.TTL_SECONDS * 2):
        messages.error(request, "この受け渡しはもう使われています。ログインしてください。")
        return redirect(f"/manage/login/?next={quote(next_url)}")
    people = User.objects.filter(is_active=True, groups__name=中["role"]).order_by("id")
    if 中["role"] == OWNER:
        people = people | User.objects.filter(is_active=True, is_superuser=True)
    user = people.distinct().first()
    if not user:
        messages.error(request, f"{中['role']} のアカウントがまだありません。パスコードを決めてください。")
        return redirect(f"/manage/login/?next={quote(next_url)}")
    login(request, user)
    return redirect(next_url)
