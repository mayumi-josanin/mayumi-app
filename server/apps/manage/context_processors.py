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


def nav(request):
    """左メニューのまとまりと、いまの画面の上部タブ（apps/manage/navigation.py）。"""
    from . import navigation

    try:
        return navigation.組み立てる(request)
    except Exception:  # 逆引きできないときも画面は出す
        return {"nav_sections": [], "nav_tabs": None, "nav_group": ""}
