from django.apps import AppConfig


class HpConfig(AppConfig):
    """公式サイト（mayumijosanin.com）の管理。

    サイトの元データと作り方は mayumi-site リポジトリの admin/ にある（core.py など）。
    ここはそれを**そのまま呼ぶ**だけで、作り方を二重に持たない（内容を同じに保つため）。
    """

    name = "apps.hp"
    label = "hp"
    verbose_name = "公式サイト"
