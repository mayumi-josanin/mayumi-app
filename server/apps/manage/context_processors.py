"""画面に「誰として入っているか」を渡す。メニューの出し分けに使う。"""

from .permissions import is_owner, roles_of


def roles(request):
    user = getattr(request, "user", None)
    if user is None:
        return {}
    return {"my_roles": sorted(roles_of(user)), "is_owner": is_owner(user)}
