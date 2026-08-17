from django.apps import AppConfig


class RecordsConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "apps.records"
    verbose_name = "院の記録（売上・仕入・操作履歴）"
