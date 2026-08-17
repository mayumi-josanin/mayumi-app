"""回数券分析のJSONを取り込む。

    python manage.py 回数券分析を取り込む 回数券分析.json --下見
    python manage.py 回数券分析を取り込む 回数券分析.json

**シートの行番号で突き合わせるので、何度実行しても二重に増えない。**

**お名前で会員に結びつけない。**この表には会員番号が無い。
CLAUDE.md にあるとおり、お名前で探すと別の方に行き着く。
一致する件数だけを数えて出し、結びつけは行わない。
"""

import json
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.members.models import Member
from apps.records.models import TicketAnalysis


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


def 写真を読む(生):
    """写真JSONを読める形にする。

    **読めなかったら捨てずに元の文字列を返す。**
    写真の場所が分からなくなると、あとから辿れない。

    返り値: (読めた値 or None, 読めなかった元の文字列)
    """
    s = 文字(生)
    if not s:
        return None, ""
    try:
        return json.loads(s), ""
    except ValueError:
        return None, s


class Command(BaseCommand):
    help = "回数券分析のJSONを取り込む"

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

        行 = 生.get("ticket_analyses") or []
        下見 = options["preview"]

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        採用, 写真が読めない = [], []

        for r in 行:
            row = 数(r.get("row"))
            回答 = 文字(r.get("response_id"))
            名 = 文字(r.get("customer_name"))
            if not row or (not 回答 and not 名):
                飛ばした += 1
                continue
            採用.append(r)

            前, 前生 = 写真を読む(r.get("before_photos"))
            後, 後生 = 写真を読む(r.get("after_photos"))
            生写真 = "\n".join(x for x in (前生, 後生) if x)
            if 生写真:
                写真が読めない.append(row)

            値 = {
                # **お名前から会員IDを埋めない。**空のまま置く。
                "member_id": "",
                "customer_name": 名[:255],
                "response_id": 回答[:128],
                "submitted_at": 日時(r.get("submitted_at")),
                "before_photos": 前,
                "after_photos": 後,
                "photos_raw": 生写真,
                "status": 文字(r.get("status"))[:64],
                "result": 文字(r.get("result")),
                "analyzed_at": 日時(r.get("analyzed_at")),
                "error": 文字(r.get("error")),
                "created_at": 日時(r.get("created_at")),
                "updated_at": 日時(r.get("updated_at")),
            }

            既存 = TicketAnalysis.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 回数券分析: JSONに {len(行)}件"
                          + (f"（シートは{生['シートの行数']}行）" if 生.get("シートの行数") else ""))
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号が無い、回答IDもお名前も空）: {飛ばした}件")

        状態 = {}
        for r in 採用:
            s = 文字(r.get("status")) or "（空）"
            状態[s] = 状態.get(s, 0) + 1
        self.stdout.write("    分析状態: " + " / ".join(
            f"{k} {n}件" for k, n in sorted(状態.items(), key=lambda x: -x[1])))

        self.stdout.write(
            f"    分析結果がある: {sum(1 for r in 採用 if 文字(r.get('result')))}件"
            f" / エラーがある: {sum(1 for r in 採用 if 文字(r.get('error')))}件")
        self.stdout.write(
            f"    ビフォー写真がある: {sum(1 for r in 採用 if 文字(r.get('before_photos')))}件"
            f" / アフター写真がある: {sum(1 for r in 採用 if 文字(r.get('after_photos')))}件")
        self.stdout.write(
            f"    回答IDが空: {sum(1 for r in 採用 if not 文字(r.get('response_id')))}件")

        if 写真が読めない:
            self.stdout.write("")
            self.stdout.write(f"    写真の記録が読めなかった行: {len(写真が読めない)}件"
                              "（元の文字列のまま残します。捨てません）")
            self.stdout.write(f"      行: {写真が読めない[:10]}")

        # **お名前の照合は数えるだけ。結びつけない。**
        名前ら = [文字(r.get("customer_name")) for r in 採用]
        名前ら = [n for n in 名前ら if n]
        一致, 複数, 無し = 0, [], []
        for n in 名前ら:
            c = Member.objects.filter(name=n).count()
            if c == 1:
                一致 += 1
            elif c > 1:
                複数.append(f"{n}（会員に{c}名）")
            else:
                無し.append(n)
        self.stdout.write("")
        self.stdout.write("  ● 会員のお名前との一致（**数えるだけ。結びつけていません**）:")
        self.stdout.write(f"      ちょうど1名と一致: {一致}件")
        if 複数:
            self.stdout.write(f"      **同じお名前が複数: {'・'.join(sorted(set(複数)))}**")
        if 無し:
            self.stdout.write(f"      一致する会員が無い: {len(無し)}件"
                              f"（{'・'.join(sorted(set(無し))[:5])}"
                              f"{'…' if len(set(無し)) > 5 else ''}）")
        self.stdout.write("      → お名前で結びつけると別の方に行き着くことがあります。")
        self.stdout.write("        結びつけるなら、お名前ではなく回答IDから辿ってください。")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for row, 値 in 新規:
                TicketAnalysis.objects.create(sheet_row=row, **値)
            for row, 値 in 更新:
                TicketAnalysis.objects.filter(sheet_row=row).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {TicketAnalysis.objects.count()}件")
