"""アプリの設定（スクリプトプロパティ）のJSONを取り込む。

    python manage.py 設定を取り込む 設定.json --下見
    python manage.py 設定を取り込む 設定.json

**鍵で突き合わせるので、何度実行しても二重に増えない。**

移すのは2つだけ。書き出し側で名指ししてある。
`ADMIN_TOKEN_SECRET` などの秘密は最初から書き出されていない。
"""

import json

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from apps.records.models import AppSetting

# 移すと決めたもの。**これ以外は取り込まない。**
# 書き出し側で絞ってあるが、受け取る側でも守る。
# 書き出しの作りが変わったときに、秘密が黙って入り込まないため。
移すもの = {
    "APP_RUNTIME_CONFIG": "アプリの版・更新案内の文言",
    "REWARD_GACHA_CONFIG": "月ごとの景品と確率",
}


class Command(BaseCommand):
    help = "アプリの設定のJSONを取り込む"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview",
                            help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        try:
            with open(options["json_path"], encoding="utf-8") as f:
                生 = json.load(f)
        except (OSError, ValueError) as e:
            raise CommandError(f"読み込めませんでした: {e}")

        設定 = 生.get("settings") or {}
        下見 = options["preview"]

        新規, 更新, 変化なし, 断った = [], [], 0, []

        for 名, 中 in 設定.items():
            # GAS が実際に返している形は、突き合わせ用に入っているだけ。
            # **取り込まない。**
            if 名.startswith("actual"):
                continue
            if not isinstance(中, dict):
                continue
            鍵 = str(中.get("key") or "").strip()
            if 鍵 not in 移すもの:
                断った.append(鍵 or 名)
                continue

            値 = 中.get("value")
            if 値 is None:
                self.stdout.write(f"    **{鍵} の中身がありません。飛ばします。**")
                continue

            既存 = AppSetting.objects.filter(key=鍵).first()
            if not 既存:
                新規.append((鍵, 値))
            elif 既存.value != 値:
                更新.append((鍵, 値, 既存.value))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write("■ アプリの設定")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 断った:
            self.stdout.write(f"    **移す対象でないので断った: {'・'.join(断った)}**")

        for 鍵, 値 in 新規:
            self.stdout.write("")
            self.stdout.write(f"  ● {鍵}（{移すもの[鍵]}）")
            self._中身を出す(値)

        for 鍵, 値, 前 in 更新:
            self.stdout.write("")
            self.stdout.write(f"  ● {鍵}（{移すもの[鍵]}）**中身が変わります**")
            self.stdout.write("      いま:")
            self._中身を出す(前, "        ")
            self.stdout.write("      これから:")
            self._中身を出す(値, "        ")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for 鍵, 値 in 新規:
                AppSetting.objects.create(key=鍵, value=値, note=移すもの[鍵])
            for 鍵, 値, _ in 更新:
                AppSetting.objects.filter(key=鍵).update(value=値, note=移すもの[鍵])

        self.stdout.write("")
        self.stdout.write(f"    → いま {AppSetting.objects.count()}件")

    def _中身を出す(self, 値, 字下げ="      "):
        """設定の中身を、読める形で並べる。長いものは要約する。"""
        if not isinstance(値, dict):
            self.stdout.write(f"{字下げ}{値}")
            return
        for k, v in 値.items():
            if isinstance(v, list):
                self.stdout.write(f"{字下げ}{k}: {len(v)}件")
                # 景品表は月ごとなので、月だけ並べる
                for x in v[:12]:
                    if isinstance(x, dict) and "month" in x:
                        等級 = len((x.get("prizes") or {}))
                        self.stdout.write(f"{字下げ}  {x['month']}（{等級}等級）")
            elif isinstance(v, dict):
                self.stdout.write(f"{字下げ}{k}: {len(v)}項目")
            else:
                t = str(v)
                if len(t) > 46:
                    t = t[:46] + "…"
                self.stdout.write(f"{字下げ}{k} = {t}")
