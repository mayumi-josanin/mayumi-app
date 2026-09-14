"""誰が何を触れるか。予約システム（mayumi-reserve/apps/core/permissions.py）と同じ考え方。

役割は **まゆみ** と **スタッフ** の2つだけ。ログインIDは持たず、
役割を選んでパスコードで入る（views_login.py）。
"""

from functools import wraps

from django.contrib.auth.models import Group
from django.core.exceptions import PermissionDenied
from django.shortcuts import redirect

OWNER = "まゆみ"
STAFF = "スタッフ"
ROLES = [OWNER, STAFF]


def ensure_groups():
    for name in ROLES:
        Group.objects.get_or_create(name=name)


def roles_of(user) -> set[str]:
    if user is None or not user.is_authenticated:
        return set()
    names = set(user.groups.values_list("name", flat=True))
    # まゆみは superuser でも入れる。役割の付け外しを間違えても締め出さないため。
    if user.is_superuser:
        names.add(OWNER)
    return names & set(ROLES)


def is_owner(user) -> bool:
    return OWNER in roles_of(user)


def owner_required(view):
    """まゆみだけの画面。

    入っていなければログインへ。入っているが役割が違えば 403。
    403 をログインへ戻すと「パスコードが違う」と思われて何度も打ち直される。
    """

    @wraps(view)
    def wrapped(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect(f"/manage/login/?next={request.path}")
        if not is_owner(request.user):
            raise PermissionDenied
        return view(request, *args, **kwargs)

    return wrapped
