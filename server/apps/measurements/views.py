"""GAS から呼ばれる読み取り専用のAPI。

この段階では**書き込みを一切受け付けない**。
スプレッドシートが正のままで、DBはまだ写しにすぎないため、
うっかり書けてしまうと「どちらが正しいか」が崩れる。
"""

import hmac

from django.conf import settings
from django.http import JsonResponse

from .models import Measurement


def _認証する(request):
    """合鍵を確かめる。合っていなければエラーの中身を返す。

    比較に hmac.compare_digest を使うのは、
    かかった時間から鍵を少しずつ言い当てられるのを防ぐため。
    """
    expected = settings.API_KEY
    if not expected:
        # 設定し忘れたまま公開してしまう事故を防ぐ。素通しにはしない。
        return JsonResponse(
            {"status": "error", "message": "APIキーが未設定です。"}, status=503
        )
    given = request.headers.get("X-Api-Key", "")
    if not hmac.compare_digest(given, expected):
        return JsonResponse(
            {"status": "error", "message": "権限がありません。"}, status=403
        )
    return None


def health(request):
    """つながっているかだけを見る。鍵は要らない。

    PCが落ちていないかを外から確かめるのに使う。
    中身は返さないので、公開されていても差し支えない。
    """
    return JsonResponse({"status": "ok"})


def measurements(request):
    """計測記録を返す。

    ?customerName=... で1人分に絞れる。
    指定が無ければ全件（いまは42件しかない）。
    """
    if request.method != "GET":
        return JsonResponse(
            {"status": "error", "message": "読み取りだけを受け付けます。"}, status=405
        )

    拒否 = _認証する(request)
    if 拒否 is not None:
        return 拒否

    query = Measurement.objects.all()

    name = (request.GET.get("customerName") or "").strip()
    if name:
        query = query.filter(customer_name=name)

    member_id = (request.GET.get("memberId") or "").strip()
    if member_id:
        query = query.filter(member_id=member_id)

    rows = [m.to_dict() for m in query]
    return JsonResponse({"status": "ok", "count": len(rows), "measurements": rows})
