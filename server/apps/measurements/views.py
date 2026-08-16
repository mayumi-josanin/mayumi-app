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
    # バイト列にしてから比べる。compare_digest は ASCII 以外の文字が入った
    # 文字列を渡すと例外になり、でたらめな合鍵を送られただけで 500 になる。
    given = request.headers.get("X-Api-Key", "")
    if not hmac.compare_digest(given.encode("utf-8", "ignore"), expected.encode("utf-8")):
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


def 入口の見え方(request):
    """呼んできた相手がどう見えているかを返す。**合鍵が要る。**

    確かめたいことが2つある。

      1. 呼び出し元のIPが1人ずつ見分けられるか。
         Funnel は 127.0.0.1 へ渡してくるので、X-Forwarded-For が
         付いていないと**全員が同じIP**に見え、呼び出し制限が全体に効いてしまう。
      2. 443番をKEM_DDENKIと分け合う形にしたとき、道（パス）がどう届くか。
         手前で削られるのか、そのまま来るのかで、URLの組み立てが変わる。

    どちらも「やってみないと分からない」ため、実物で見るための窓口。
    合鍵を要るようにしてあるのは、中の作りを外から覗かせないため。
    """
    拒否 = _認証する(request)
    if 拒否 is not None:
        return 拒否

    from apps.core.middleware import 呼び出し元のIP

    return JsonResponse(
        {
            "status": "ok",
            "見えているIP": 呼び出し元のIP(request),
            "REMOTE_ADDR": request.META.get("REMOTE_ADDR", ""),
            "X-Forwarded-For": request.META.get("HTTP_X_FORWARDED_FOR", ""),
            "X-Forwarded-Proto": request.META.get("HTTP_X_FORWARDED_PROTO", ""),
            "届いた道": request.path,
            "届いた道（全体）": request.get_full_path(),
            "Host": request.get_host(),
            "Origin": request.META.get("HTTP_ORIGIN", ""),
        }
    )


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
