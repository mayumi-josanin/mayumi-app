"""ビジリスの表の正を切り替える印を立てる／下ろす。

    python manage.py ビジリスの正を切り替える server   … 9/24 以降、ビジリスの Code.gs に転送を入れて deploy したあと
    python manage.py ビジリスの正を切り替える sheet    … 戻したとき
    python manage.py ビジリスの正を切り替える          … いまの印を見るだけ

管理画面のビジリスまわりの書き込みは、この印が server のときだけ通る（apps/bijiris/gate.py）。
"""

from django.core.management.base import BaseCommand, CommandError

from apps.bijiris import gate


class Command(BaseCommand):
    help = "ビジリスの表の正（server / sheet）を切り替える印"

    def add_arguments(self, parser):
        parser.add_argument("先", nargs="?", choices=["server", "sheet"])

    def handle(self, *args, **opts):
        if opts["先"]:
            gate.切り替える(opts["先"])
        いま = "server" if gate.ビジリスはサーバーが正() else "sheet"
        self.stdout.write(f"bijiris_source = {いま}")
        if opts["先"] and opts["先"] != いま:
            raise CommandError("印が立ちませんでした")
