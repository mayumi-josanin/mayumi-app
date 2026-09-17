from django.apps import AppConfig


class PreviewConfig(AppConfig):
    """お客様の画面（確認用）。

    院長の依頼（2026-09-18）:「お客様アプリがあるやつ（ビジリス・まゆみ助産院お客様アプリ・予約システム）を
    develop 環境で確認して見れるようにできますか」。3つを1か所（左メニューの「お客様の画面」）にまとめる。
    読むだけの画面で、保存するものが無いので models は持たない。
    """

    default_auto_field = "django.db.models.BigAutoField"
    name = "apps.preview"
    label = "preview"
    verbose_name = "お客様の画面"
