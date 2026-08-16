"""アプリから直接呼ばれるための、入口の守り。

いまは GAS だけが呼ぶので `API_KEY`（サーバー同士の共通鍵）で足りている。
アプリが直接呼ぶようになると、次の2つが要る。

  CORS … アプリは GitHub Pages（mayumi-josanin.github.io）で配っており、
         APIとは住所が違う。ブラウザは既定で住所をまたぐ通信を止めるので、
         「この住所からは受ける」と明示する。

  制限 … 公開の窓口が自宅PCに直接向くため、同じ相手から短時間に大量に
         来ると、家庭の回線とPCではすぐに詰まる。断る仕組みを入れる。

**どちらも外部の部品を足していない。**自宅PCに置くので、
入れ替えの要る部品を増やしたくない（更新を忘れたまま何年も動くため）。
"""

import time

from django.core.cache import cache
from django.http import HttpResponse, JsonResponse


def 呼び出し元のIP(request):
    """呼んできた相手のIPを取る。

    Tailscale Funnel は TLS を終端して 127.0.0.1 へ渡してくるので、
    そのままだと**全員が同じIP**に見えて、制限が全体に効いてしまう。
    Funnel が付けてくれる X-Forwarded-For を先に見る。

    偽装が心配になるが、Django は 127.0.0.1 でしか待ち受けていない
    （docker-compose.yml でそう縛ってある）。つまりこのヘッダを付けられるのは
    Funnel だけで、外から直接は届かない。
    """
    転送元 = request.META.get("HTTP_X_FORWARDED_FOR", "")
    if 転送元:
        # 一番左が元の相手。あいだに挟まったものが右へ並ぶ。
        return 転送元.split(",")[0].strip()
    return request.META.get("REMOTE_ADDR", "") or "unknown"


class CORSミドルウェア:
    """許した住所からの通信にだけ、ブラウザ向けの許可を付ける。

    `*`（どこからでも）にはしない。それをすると、どこかの見知らぬページの
    JavaScript からも、お客様の合鍵さえあれば読めてしまう。
    """

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        from django.conf import settings

        許した = getattr(settings, "CORS_ALLOWED_ORIGINS", [])
        元 = request.META.get("HTTP_ORIGIN", "")
        通す = 元 in 許した

        # 下見の問い合わせ（プリフライト）。本体を動かさずにここで返す。
        # ブラウザは、独自ヘッダを付ける通信の前に必ずこれを送ってくる。
        if request.method == "OPTIONS" and request.META.get("HTTP_ACCESS_CONTROL_REQUEST_METHOD"):
            response = HttpResponse(status=204 if 通す else 403)
        else:
            response = self.get_response(request)

        # 許していない住所でも Vary は必ず付ける。
        # 付けないと、許した住所への応答が中継に覚えられ、
        # 別の住所へそのまま返ってしまうことがある。
        response["Vary"] = self._varyに足す(response.get("Vary", ""), "Origin")

        if 通す:
            response["Access-Control-Allow-Origin"] = 元
            response["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
            response["Access-Control-Allow-Headers"] = "Content-Type, X-Api-Key, Authorization"
            # 下見の結果を1日覚えてもらう。毎回2往復になるのを避ける。
            response["Access-Control-Max-Age"] = "86400"

        return response

    @staticmethod
    def _varyに足す(いまの値, 足すもの):
        既存 = [v.strip() for v in str(いまの値 or "").split(",") if v.strip()]
        if 足すもの.lower() not in [v.lower() for v in 既存]:
            既存.append(足すもの)
        return ", ".join(既存)


class 管理画面をしまうミドルウェア:
    """Django の管理画面を、外から開けないようにする。

    Funnel は「この口に来たものは全部通す」という作りなので、
    `/admin/` も一緒にインターネットへ出てしまう。会員情報を扱う画面が
    ログイン欄ごと外に見えているのは、鍵を掛けていても好ましくない。

    このPCの中から開いたときだけ通す。手元から見たいときは、
    SSHの転送でこのPCの中を経由して開く（外には開かない）。

        ssh -L 8002:127.0.0.1:8002 （ユーザー名）@desktop-rmsk0vg.tail8efe0d.ts.net
        → 手元のブラウザで http://localhost:8002/admin/

    403 ではなく 404 を返すのは、そこに管理画面があること自体を
    外に教えないため。
    """

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        from django.conf import settings

        if request.path.startswith("/admin"):
            住所 = request.get_host().split(":")[0]
            if 住所 not in getattr(settings, "ADMIN_ALLOWED_HOSTS", []):
                return HttpResponse("Not Found", status=404, content_type="text/plain")

        return self.get_response(request)


class 呼び出し制限ミドルウェア:
    """同じ相手からの呼び出しが多すぎたら断る。

    1分ごとの区切りで数える。区切りをまたぐと0に戻る素朴な作りだが、
    「家庭の回線とPCを守る」目的にはこれで足りる。

    数を覚える場所はデータベース側の入れ物にしてある（`createcachetable`）。
    Django の既定はプロセスごとに別々に覚えるので、
    働き手（gunicorn worker）が複数あると、実際には人数ぶん通ってしまう。
    """

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        from django.conf import settings

        上限, 区切り秒 = self._この道の上限(request, settings)
        if 上限 <= 0:
            return self.get_response(request)

        ip = 呼び出し元のIP(request)
        窓 = int(time.time() // 区切り秒)
        鍵 = f"throttle:{窓}:{区切り秒}:{ip}"

        try:
            # add は「無いときだけ置く」。すでにあれば何もしない。
            cache.add(鍵, 0, 区切り秒 + 5)
            回数 = cache.incr(鍵)
        except ValueError:
            # 数え始めた直後に区切りが変わると、鍵が消えていることがある。
            cache.set(鍵, 1, 区切り秒 + 5)
            回数 = 1
        except Exception:
            # 数える入れ物が使えない（createcachetable のし忘れなど）。
            # **ここで止めない。**数えられないことを理由に全部断ると、
            # 制限のための仕組みが、そのままアプリの全停止になってしまう。
            # 数を数えるのは守りの足しであって、無くても本来の鍵の確認は生きている。
            return self.get_response(request)

        if 回数 > 上限:
            残り = 区切り秒 - int(time.time() % 区切り秒)
            response = JsonResponse(
                {
                    "status": "error",
                    "message": "呼び出しが多すぎます。少し時間をおいてからお試しください。",
                },
                status=429,
            )
            response["Retry-After"] = str(max(1, 残り))
            return response

        return self.get_response(request)

    @staticmethod
    def _この道の上限(request, settings):
        道 = request.path

        # 生死の確認は数えない。見張りが断られると、
        # 「止まっている」と誤って知らせてしまう。
        for 除く in getattr(settings, "THROTTLE_EXEMPT_PATHS", []):
            if 道.startswith(除く):
                return 0, 60

        # ログインなど、当てずっぽうを繰り返される道は厳しくする。
        for 厳しい in getattr(settings, "THROTTLE_STRICT_PATHS", []):
            if 道.startswith(厳しい):
                return getattr(settings, "THROTTLE_STRICT_PER_MINUTE", 10), 60

        return getattr(settings, "THROTTLE_PER_MINUTE", 120), 60
