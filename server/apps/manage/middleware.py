"""ログインの総当たりを待たせる。予約システムの LoginThrottleMiddleware と同じ。

パスコードは4桁の数字なので、止めなければ1万回で必ず当たる。
**失敗した POST だけ**を数え、15分に5回を超えたら待っていただく。
成功や画面を開くだけの GET は数えない。
"""

from django.core.cache import cache
from django.http import HttpResponse

WINDOW = 15 * 60
LIMIT = 5


def _key(request) -> str:
    ip = request.META.get("HTTP_X_FORWARDED_FOR", "").split(",")[0].strip() or request.META.get(
        "REMOTE_ADDR", ""
    )
    role = request.POST.get("role", "")
    return f"manage-login-fail:{ip}:{role}"


class ログイン制限ミドルウェア:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        対象 = request.method == "POST" and request.path.endswith("/login/")
        if 対象:
            try:
                if int(cache.get(_key(request)) or 0) >= LIMIT:
                    return HttpResponse(
                        "パスコードの間違いが続いたため、15分ほど待ってからお試しください。",
                        status=429,
                        content_type="text/plain; charset=utf-8",
                    )
            except Exception:
                pass
        response = self.get_response(request)
        if 対象:
            try:
                # ログインに成功すると次の画面へ飛ぶ（302）。それ以外は失敗。
                if response.status_code != 302:
                    鍵 = _key(request)
                    cache.add(鍵, 0, WINDOW)
                    try:
                        cache.incr(鍵)
                    except ValueError:
                        cache.set(鍵, 1, WINDOW)
            except Exception:
                pass
        return response
