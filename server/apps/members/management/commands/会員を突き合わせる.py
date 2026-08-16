"""スプレッドシートの中身と、会員台帳を1件ずつ突き合わせる。

    python manage.py 会員を突き合わせる 会員データ.json

**読むだけ。**何も書かない。

取り込みが「終わった」ことと「正しく移った」ことは別。
件数が合っていても、中身がずれていれば意味がない。
移し終わったら必ずこれを通し、**食い違いがゼロ**を目で見てから次へ進む。

パスコードは、平文とハッシュを照合して確かめる。
「ログインできなくなっていないか」がいちばん怖いところなので、
数え上げではなく実際に照合する。
"""

import json

from django.contrib.auth.hashers import check_password
from django.core.management.base import BaseCommand, CommandError

from apps.members.models import Member


class Command(BaseCommand):
    help = "スプレッドシートのJSONと会員台帳を突き合わせる（読むだけ）"

    def add_arguments(self, parser):
        parser.add_argument("json_path")

    def handle(self, *args, **options):
        try:
            with open(options["json_path"], encoding="utf-8") as f:
                生 = json.load(f)
        except (OSError, ValueError) as e:
            raise CommandError(f"読み込めませんでした: {e}")

        行 = 生.get("members") if isinstance(生, dict) else 生
        シート = {str(r.get("memberId") or "").strip(): r for r in 行 if str(r.get("memberId") or "").strip()}
        台帳 = {m.member_id: m for m in Member.objects.all()}

        シートだけ = sorted(set(シート) - set(台帳))
        台帳だけ = sorted(set(台帳) - set(シート))
        両方 = sorted(set(シート) & set(台帳))

        食い違い = []
        パスコード合致, パスコード不一致, パスコード無し = 0, [], 0

        for mid in 両方:
            s, m = シート[mid], 台帳[mid]

            def 比べる(項目, シート値, 台帳値):
                a = str(シート値 or "").strip()
                b = str(台帳値 or "").strip()
                if a != b:
                    食い違い.append(f"{mid} {m.name} / {項目}: シート「{a}」 台帳「{b}」")

            比べる("氏名", s.get("name"), m.name)
            比べる("フリガナ", s.get("kana"), m.kana)
            比べる("電話番号", s.get("phone"), m.phone)
            比べる("住所", s.get("address"), m.address)
            比べる("生年月日", s.get("birthday"), m.birthday.isoformat() if m.birthday else "")

            シートのスタンプ = int(float(s.get("stampCount") or 0))
            if シートのスタンプ != m.stamp_count:
                食い違い.append(f"{mid} {m.name} / スタンプ数: シート{シートのスタンプ} 台帳{m.stamp_count}")

            平文 = str(s.get("passcode") or "").strip()
            if not 平文:
                パスコード無し += 1
            elif m.passcode_hash and check_password(平文, m.passcode_hash):
                パスコード合致 += 1
            else:
                パスコード不一致.append(f"{mid} {m.name}")

        self.stdout.write("")
        self.stdout.write(f"■ シート {len(シート)}名 ／ 台帳 {len(台帳)}名")
        self.stdout.write("")

        if シートだけ:
            self.stdout.write(f"■ **シートにあって台帳に無い方: {len(シートだけ)}名**")
            for x in シートだけ[:20]:
                self.stdout.write(f"    {x} {シート[x].get('name', '')}")
            self.stdout.write("")
        if 台帳だけ:
            self.stdout.write(f"■ 台帳にあってシートに無い方: {len(台帳だけ)}名")
            self.stdout.write("    （シートから消された方。台帳には残ります）")
            for x in 台帳だけ[:20]:
                self.stdout.write(f"    {x} {台帳[x].name}")
            self.stdout.write("")

        self.stdout.write(f"■ 中身の食い違い: {len(食い違い)}件")
        for x in 食い違い[:30]:
            self.stdout.write(f"    {x}")
        if len(食い違い) > 30:
            self.stdout.write(f"    …ほか {len(食い違い) - 30}件")
        self.stdout.write("")

        self.stdout.write("■ パスコードで実際にログインできるか")
        self.stdout.write(f"    照合できた:       {パスコード合致}名")
        self.stdout.write(f"    シートに未設定:   {パスコード無し}名")
        if パスコード不一致:
            self.stdout.write(f"    **照合できない:   {len(パスコード不一致)}名**")
            for x in パスコード不一致[:20]:
                self.stdout.write(f"      {x}")
        self.stdout.write("")

        問題 = len(シートだけ) + len(食い違い) + len(パスコード不一致)
        if 問題 == 0:
            self.stdout.write("■ 食い違いはありません。次へ進めます。")
        else:
            self.stdout.write(f"■ **{問題}件の食い違いがあります。直してから次へ進んでください。**")
