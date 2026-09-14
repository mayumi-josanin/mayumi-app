"""管理画面のパスコードを決める（まゆみ／スタッフ）。

    MANAGE_PASSCODE=1234 python manage.py manage_passcode まゆみ

パスコードはコマンドの引数に書かない（履歴に残る）。環境変数で渡す。
無ければ何もしない。**アカウントは役割名そのもの**（まゆみ／スタッフ）で、
無ければ作り、その役割のグループに入れる。
"""

import os
import unicodedata

from django.contrib.auth.models import Group, User
from django.core.management.base import BaseCommand, CommandError

from apps.manage.permissions import OWNER, ROLES, STAFF, ensure_groups

# PowerShell 越しに日本語を渡すと壊れるので、英字の呼び名も受ける。
ALIASES = {"owner": OWNER, "staff": STAFF}


class Command(BaseCommand):
    help = "管理画面のパスコードを決めます（環境変数 MANAGE_PASSCODE から読みます）"

    def add_arguments(self, parser):
        parser.add_argument("role", choices=ROLES + list(ALIASES), help="まゆみ(owner) か スタッフ(staff)")

    def handle(self, *args, **options):
        role = ALIASES.get(options["role"], options["role"])
        passcode = unicodedata.normalize("NFKC", os.environ.get("MANAGE_PASSCODE", "")).strip()
        if not passcode:
            raise CommandError("環境変数 MANAGE_PASSCODE が空です。パスコードを入れて実行してください。")
        if len(passcode) < 4:
            raise CommandError("パスコードは4桁以上にしてください。")

        ensure_groups()
        user, created = User.objects.get_or_create(username=role)
        user.set_password(passcode)
        user.is_active = True
        user.save()
        user.groups.add(Group.objects.get(name=role))
        self.stdout.write(
            self.style.SUCCESS(f"{role} のパスコードを{'作りました' if created else '変えました'}。")
        )
