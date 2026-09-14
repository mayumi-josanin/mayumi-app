from django.apps import AppConfig


class ManageConfig(AppConfig):
    """院の管理画面。

    GitHub Pages の管理アプリ（admin/index.html）を、1画面ずつここへ移す。
    GAS を通らず、この Django が直接データベースを読み書きする。
    お客様アプリは触らない（GAS の窓口のまま）。

    **Funnel（mayumi-api）には出さない。**`MANAGE_ENABLED` を立てた箱だけが
    この画面を持ち、その箱は Tailscale の中にしか出さない（予約システムの
    管理画面と同じ分け方）。
    """

    default_auto_field = "django.db.models.BigAutoField"
    name = "apps.manage"
    verbose_name = "管理画面"
