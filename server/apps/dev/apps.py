from django.apps import AppConfig


class DevConfig(AppConfig):
    """開発（開発管理と目安箱）。

    KEM_DDENKI（建設 SaaS）の「開発管理」の写し。プロジェクト → タスク → コメントの3段と、
    使う人からの意見を受ける目安箱。院長の依頼（2026-09-16）で、同じ画面と項目をこちらの
    管理画面にも持つ。会社（tenant）や履歴（simple_history）は無いので、その分だけ簡単になっている。
    """

    default_auto_field = "django.db.models.BigAutoField"
    name = "apps.dev"
    label = "dev"
    verbose_name = "開発"
