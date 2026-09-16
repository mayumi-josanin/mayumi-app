"""ビジリスの記録は、いまサーバーが正か（切り替えの印）。

apps/manage/member_gate.py と同じ作り。会員の切り替え（9/19〜23）のあと、9/24 以降に
ビジリスの Code.gs へ転送を入れて `clasp deploy -i` してから、この印を立てる。
**印が立つまでにサーバーのビジリスの表へ書くと、スプレッドシートには無い記録ができて正が2つになる。**
管理画面のビジリスまわりの書き込みは、この印が立つまで断る（読むのはよい。古いだけ）。

印は AppSetting の `bijiris_source`（"server" なら立っている）。
`python manage.py ビジリスの正を切り替える server`（戻すときは `sheet`）で立てる。
"""

from apps.records.models import AppSetting

鍵 = "bijiris_source"


def ビジリスはサーバーが正() -> bool:
    s = AppSetting.objects.filter(pk=鍵).first()
    return bool(s) and (s.value or {}).get("source") == "server"


def 断る文():
    return "ビジリスの記録はまだスプレッドシートが正です。切り替え（9/24 以降）までは旧ビジリス管理アプリで直してください。"


def 切り替える(先: str):
    """'server' か 'sheet'。"""
    if 先 not in ("server", "sheet"):
        raise ValueError("先は server か sheet のどちらかです")
    AppSetting.objects.update_or_create(
        pk=鍵, defaults={"value": {"source": 先}, "note": "ビジリスの表の正（server=サーバー / sheet=スプレッドシート）"})
