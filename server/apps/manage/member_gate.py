"""会員の表は、いまサーバーが正か（切り替えの印）。

9/19 の切り替えまで、会員の正はスプレッドシート。**それまでにサーバーの会員へ書くと、
シートには無い記録ができて正が2つになる。**管理画面の会員まわりの書き込みは、
この印が立つまで断る（読むのはよい。古いだけ）。

印は AppSetting の `member_source`（"server" なら立っている）。
切り替え当日に `python manage.py 会員の正を切り替える server`（戻すときは `sheet`）で立てる。
サーバーPCの補助 `~/member_cutover.sh live-on / live-off` が同じことをする。
"""

from apps.records.models import AppSetting

鍵 = "member_source"


def 会員はサーバーが正() -> bool:
    s = AppSetting.objects.filter(pk=鍵).first()
    return bool(s) and (s.value or {}).get("source") == "server"


def 断る文():
    return "会員の記録はまだスプレッドシートが正です。切り替え（9/19）までは旧管理アプリで直してください。"


def 切り替える(先: str):
    """'server' か 'sheet'。"""
    AppSetting.objects.update_or_create(
        pk=鍵, defaults={"value": {"source": 先}, "note": "会員の表の正（server=サーバー / sheet=スプレッドシート）"})
