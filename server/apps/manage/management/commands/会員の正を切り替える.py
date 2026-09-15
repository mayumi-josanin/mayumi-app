"""会員の表の正を切り替える印を立てる／下ろす。

    python manage.py 会員の正を切り替える server   … 9/19 に SERVER_TABLES へ member を入れたあと
    python manage.py 会員の正を切り替える sheet    … 戻したとき
    python manage.py 会員の正を切り替える          … いまの印を見るだけ

管理画面の会員まわりの書き込みは、この印が server のときだけ通る（apps/manage/member_gate.py）。
"""

from django.core.management.base import BaseCommand, CommandError

from apps.manage import member_gate


class Command(BaseCommand):
    help = "会員の表の正（server / sheet）を切り替える印"

    def add_arguments(self, parser):
        parser.add_argument("先", nargs="?", choices=["server", "sheet"])

    def handle(self, *args, **opts):
        if opts["先"]:
            member_gate.切り替える(opts["先"])
        いま = "server" if member_gate.会員はサーバーが正() else "sheet"
        self.stdout.write(f"member_source = {いま}")
        if opts["先"] and opts["先"] != いま:
            raise CommandError("印が立ちませんでした")
