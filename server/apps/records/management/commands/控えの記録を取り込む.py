"""控えの記録のJSONを取り込む。

    python manage.py 控えの記録を取り込む 控えの記録.json --下見
    python manage.py 控えの記録を取り込む 控えの記録.json

**シートの行番号で突き合わせるので、何度実行しても二重に増えない。**
"""

import json
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.records.models import BackupRecord


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値):
    if 値 is None or 値 == "":
        return None
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


def 日時(値):
    s = 文字(値)
    if not s:
        return None
    try:
        d = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if timezone.is_naive(d):
        d = timezone.make_aware(d, timezone.get_current_timezone())
    return d


class Command(BaseCommand):
    help = "控えの記録のJSONを取り込む"

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

        行 = 生.get("backup_records") or []
        下見 = options["preview"]

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        採用 = []

        for r in 行:
            row = 数(r.get("row"))
            日 = 日時(r.get("created_at"))
            名 = 文字(r.get("file_name"))
            if not row or (not 日 and not 名):
                飛ばした += 1
                continue
            採用.append(r)

            値 = {
                "created_at": 日,
                "kind": 文字(r.get("kind"))[:64],
                "file_name": 名[:255],
                "file_id": 文字(r.get("file_id"))[:128],
                "url": 文字(r.get("url")),
            }

            既存 = BackupRecord.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 控えの記録: JSONに {len(行)}件"
                          + (f"（シートは{生['シートの行数']}行）" if 生.get("シートの行数") else ""))
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号が無い、日時もファイル名も空）: {飛ばした}件")

        種 = {}
        for r in 採用:
            k = 文字(r.get("kind")) or "（空）"
            種[k] = 種.get(k, 0) + 1
        self.stdout.write("    種別: " + " / ".join(
            f"{k} {n}件" for k, n in sorted(種.items(), key=lambda x: -x[1])))

        ID無し = sum(1 for r in 採用 if not 文字(r.get("file_id")))
        URL無し = sum(1 for r in 採用 if not 文字(r.get("url")))
        self.stdout.write(f"    ファイルIDが無い: {ID無し}件 / URLが無い: {URL無し}件")

        日ら = [日時(r.get("created_at")) for r in 採用]
        日ら = [d for d in 日ら if d]
        if 日ら:
            self.stdout.write(f"    期間: {min(日ら).date()} 〜 {max(日ら).date()}")
        日時無し = len(採用) - len(日ら)
        if 日時無し:
            self.stdout.write(f"    **日時が無い: {日時無し}件**")

        # 同じファイルIDが2行あると、同じ控えを二重に記録している。
        ID = {}
        for r in 採用:
            i = 文字(r.get("file_id"))
            if i:
                ID[i] = ID.get(i, 0) + 1
        重なり = {k: v for k, v in ID.items() if v > 1}
        if 重なり:
            self.stdout.write("")
            self.stdout.write(f"    **同じファイルIDが {len(重なり)}種あります"
                              f"（同じ控えを二重に記録）:**")
            for i, n in list(重なり.items())[:10]:
                self.stdout.write(f"      {i} … {n}行")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            BackupRecord.objects.bulk_create(
                [BackupRecord(sheet_row=row, **値) for row, 値 in 新規], batch_size=500)
            for row, 値 in 更新:
                BackupRecord.objects.filter(sheet_row=row).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {BackupRecord.objects.count()}件")
