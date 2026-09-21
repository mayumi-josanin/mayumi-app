"""お客様アプリの住所の既定を入れる（QRコード案内で使う）。

旧管理アプリは、この住所をそのパソコンのブラウザの中（localStorage）に持っていたので、
サーバーには何も無い。**空のままだとQRコード案内が使えない**ので、いま配っている住所を
最初から入れておく。すでに決まっているときは触らない。
"""

from django.db import migrations

鍵 = "app_public_url"
既定 = "https://mayumi-josanin.github.io/mayumi-app/"


def 入れる(apps, schema_editor):
    AppSetting = apps.get_model("records", "AppSetting")
    if not AppSetting.objects.filter(pk=鍵).exists():
        AppSetting.objects.create(key=鍵, value={"url": 既定},
                                  note="お客様アプリの住所（QRコード案内で使う）")


def 戻す(apps, schema_editor):
    apps.get_model("records", "AppSetting").objects.filter(pk=鍵).delete()


class Migration(migrations.Migration):

    dependencies = [("records", "0010_orderline")]

    operations = [migrations.RunPython(入れる, 戻す)]
