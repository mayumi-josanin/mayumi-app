from django.apps import AppConfig


class BijirisConfig(AppConfig):
    """ビジリス（EMS トレーニング）の管理。

    ビジリスの GAS（bijiris/gas/Code.gs）が持っていた表を、こちらのサーバーに写す
    （docs/design/ビジリス管理を統合する.md）。お客様のビジリスアプリは変えない。
    GAS を窓口に残し、表ごとに移して切り替えの印（gate.py）で戻せるようにする。
    """

    default_auto_field = "django.db.models.BigAutoField"
    name = "apps.bijiris"
    label = "bijiris"
    verbose_name = "ビジリス"
