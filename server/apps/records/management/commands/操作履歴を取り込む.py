"""操作履歴のJSONを取り込む。

    python manage.py 操作履歴を取り込む 操作履歴.json --下見
    python manage.py 操作履歴を取り込む 操作履歴.json

**シートの行番号で突き合わせるので、何度実行しても二重に増えない。**

移すのは「人の操作」だけ。自動処理の2種（9,373件中8,579件）は
書き出し側で除いてある。**受け取る側でも同じ規則で弾く。**
書き出しの作りが変わったときに、自動処理が黙って混ざらないため。
"""

import json
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.records.models import AuditLog


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


def 詳細を読む(生):
    """詳細JSONを読める形にする。

    **読めなかったら捨てずに、元の文字列を返す。**
    形が違うからといって記録を消さない。あとで見返したときに
    「何かが入っていたはずなのに空」では困る。

    返り値: (読めた辞書 or None, 読めなかった元の文字列)
    """
    s = 文字(生)
    if not s:
        return None, ""
    try:
        x = json.loads(s)
    except ValueError:
        return None, s
    # 配列や数値が来ることもある。辞書でなければ包んで持つ。
    if isinstance(x, dict):
        return x, ""
    return {"値": x}, ""


class Command(BaseCommand):
    help = "操作履歴のJSONを取り込む"

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

        行 = 生.get("audit_logs") or []
        下見 = options["preview"]

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        自動が混ざった = {}
        種別ごと, 読めない = {}, []
        採用 = []

        for r in 行:
            row = 数(r.get("row"))
            種 = 文字(r.get("kind"))
            日 = 日時(r.get("happened_at"))

            if not row or (not 種 and not 日):
                飛ばした += 1
                continue

            # 書き出し側で除いてあるはずだが、ここでも弾く。
            # **受け取る側で守らないと、書き出しの作りが変わったときに
            # 自動処理8,579件が黙って混ざる。**
            if 種 in AuditLog.自動処理の種別:
                自動が混ざった[種] = 自動が混ざった.get(種, 0) + 1
                continue

            採用.append(r)
            種別ごと[種] = 種別ごと.get(種, 0) + 1

            詳細, 生詳細 = 詳細を読む(r.get("detail_raw"))
            if 生詳細:
                読めない.append(row)

            値 = {
                "happened_at": 日,
                "kind": 種[:64],
                "result": 文字(r.get("result"))[:32],
                "target": 文字(r.get("target")),
                "summary": 文字(r.get("summary")),
                "operator": 文字(r.get("operator"))[:64],
                "detail": 詳細,
                "detail_raw": 生詳細,
            }

            既存 = AuditLog.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 操作履歴: JSONに {len(行)}件"
                          + (f"（シートは{生['シートの行数']}行）" if 生.get("シートの行数") else ""))
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号が無い、種別も日時も空）: {飛ばした}件")

        if 自動が混ざった:
            # 起きてはいけないこと。起きたら必ず気づけるようにする。
            self.stdout.write("")
            self.stdout.write("    **自動処理が混ざっていました。ここで弾きました:**")
            for k, n in sorted(自動が混ざった.items(), key=lambda x: -x[1]):
                self.stdout.write(f"      {k} … {n}件")

        if 読めない:
            self.stdout.write("")
            self.stdout.write(f"    詳細が読めなかった行: {len(読めない)}件"
                              f"（元の文字列のまま残します。捨てません）")
            self.stdout.write(f"      行: {読めない[:10]}"
                              + ("…" if len(読めない) > 10 else ""))

        失敗 = sum(1 for r in 採用 if 文字(r.get("result")) == "失敗")
        概要あり = sum(1 for r in 採用 if 文字(r.get("summary")))
        対象あり = sum(1 for r in 採用 if 文字(r.get("target")))
        操作者 = {文字(r.get("operator")) for r in 採用}
        self.stdout.write("")
        self.stdout.write(f"    結果が「失敗」: {失敗}件")
        self.stdout.write(f"    概要がある: {概要あり}件 / 対象がある: {対象あり}件")
        self.stdout.write(f"    操作者: {'・'.join(sorted(操作者)) or '（空）'}")

        if 採用:
            日ら = [日時(r.get("happened_at")) for r in 採用]
            日ら = [d for d in 日ら if d]
            if 日ら:
                self.stdout.write(f"    期間: {min(日ら).date()} 〜 {max(日ら).date()}")

        self.stdout.write("")
        self.stdout.write(f"  ● 種別（多い順・{len(種別ごと)}種）:")
        for k, n in sorted(種別ごと.items(), key=lambda x: -x[1]):
            self.stdout.write(f"      {k} … {n}件")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            AuditLog.objects.bulk_create(
                [AuditLog(sheet_row=row, **値) for row, 値 in 新規], batch_size=500)
            for row, 値 in 更新:
                AuditLog.objects.filter(sheet_row=row).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {AuditLog.objects.count()}件")
