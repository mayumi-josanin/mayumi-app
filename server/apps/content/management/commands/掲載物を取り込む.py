"""お知らせとカレンダーのJSONを取り込む。

    python manage.py 掲載物を取り込む 掲載物.json --下見
    python manage.py 掲載物を取り込む 掲載物.json

**シートの行番号で突き合わせるので、何度実行しても二重に増えない。**

会員のような安定した鍵が無い表なので、行番号を鍵にしている。
シートの行を挿入・削除すると番号がずれるが、移行のあいだは
シートを触らない前提で進める。正をサーバーへ移したあとは、
サーバー側のIDが鍵になるのでこの弱さは消える。
"""

import json
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.content.models import CalendarEvent, News


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値):
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


def 日付(値):
    s = 文字(値)[:10]
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
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


def 公開か(値):
    """「公開」以外は出さない扱いにする。

    空欄を公開とみなすと、書きかけのものがお客様に出てしまう。
    **迷ったら出さない**側に倒す。
    """
    return 文字(値) == "公開"


def 削除済みか(値):
    return bool(文字(値))


class Command(BaseCommand):
    help = "お知らせとカレンダーのJSONを取り込む"

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

        下見 = options["preview"]
        結果 = []
        結果.append(self._取り込む(News, 生.get("news") or [], "お知らせ", 下見))
        結果.append(self._取り込む(CalendarEvent, 生.get("calendar") or [], "カレンダー", 下見))

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")

    def _取り込む(self, model, 行, 名, 下見):
        日付の鍵 = "posted_on" if model is News else "event_on"
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0

        for r in 行:
            row = 数(r.get("row"))
            if not row:
                飛ばした += 1
                continue

            # 題も日付も無い行は、書きかけの空行。書き出し側でも弾いているが、
            # ここでも弾く。**受け取る側で守らないと、書き出しの作りが
            # 変わったときに空行が黙って入り込む。**
            if not 文字(r.get("title")) and not 日付(r.get(日付の鍵)):
                飛ばした += 1
                continue

            値 = {
                "title": 文字(r.get("title"))[:255],
                "published": 公開か(r.get("published")),
                "notice_listed": 公開か(r.get("notice_listed")),
                "deleted": 削除済みか(r.get("deleted")),
                "deleted_at": 日時(r.get("deleted_at")),
                "delete_reason": 文字(r.get("delete_reason"))[:255],
                "publish_at": 日時(r.get("publish_at")),
                "notice_listed_at": 日時(r.get("notice_listed_at")),
                "notice_delisted_at": 日時(r.get("notice_delisted_at")),
                "sort_order": 数(r.get("sort_order")),
                "image_url": 文字(r.get("image_url")),
                "updated_at": 日時(r.get("updated_at")),
            }
            if model is News:
                値.update({
                    "posted_on": 日付(r.get("posted_on")),
                    "category": 文字(r.get("category"))[:100],
                    "icon": 文字(r.get("icon"))[:16],
                    "body": 文字(r.get("body")),
                    "link_url": 文字(r.get("link_url")),
                    "link_label": 文字(r.get("link_label"))[:100],
                    "button_text": 文字(r.get("button_text"))[:100],
                })
            else:
                値.update({
                    "event_on": 日付(r.get("event_on")),
                    "detail": 文字(r.get("detail")),
                    "color": 文字(r.get("color"))[:32],
                    "menu_row": 数(r.get("menu_row")),
                })

            既存 = model.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
                continue
            変わった = [k for k, v in 値.items() if getattr(既存, k) != v]
            if 変わった:
                更新.append((row, 値, 変わった))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ {名}: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号が無い・題も日付も空）: {飛ばした}件")

        公開数 = sum(1 for r in 行 if 公開か(r.get("published")))
        削除数 = sum(1 for r in 行 if 削除済みか(r.get("deleted")))
        self.stdout.write(f"    うち公開: {公開数}件 / 削除済み: {削除数}件")

        if 下見:
            return

        with transaction.atomic():
            for row, 値 in 新規:
                model.objects.create(sheet_row=row, **値)
            for row, 値, _ in 更新:
                model.objects.filter(sheet_row=row).update(**値)

        self.stdout.write(f"    → いま {model.objects.count()}件")
