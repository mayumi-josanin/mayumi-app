"""画面に「誰として入っているか」を渡す。メニューの出し分けに使う。"""

from django.conf import settings

from .permissions import is_owner, roles_of


def roles(request):
    user = getattr(request, "user", None)
    if user is None:
        return {}
    return {"my_roles": sorted(roles_of(user)), "is_owner": is_owner(user)}


def static_version(request):
    """CSS / JS の版（?v=）。箱を作り直すたびに変わるので、端末の古い控えが使われない。"""
    return {"STATIC_VERSION": getattr(settings, "STATIC_VERSION", "")}
