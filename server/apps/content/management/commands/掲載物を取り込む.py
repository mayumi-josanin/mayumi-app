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
import re
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.content.models import (
    CalendarEvent,
    Category,
    Menu,
    News,
    Product,
    PushNotice,
    SupportFaq,
)


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値):
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


# 読めなかった日付を覚えておく。**黙って捨てない。**
# メニューの登録日13件が、読めない書き方（`Mon Mar 30 2026 …`）で来ていて
# 全部 None になっていた。取り込みは成功したように見えていた（2026-09-05）。
読めなかった日付 = []


def 日付(値):
    元 = 文字(値)
    if not 元:
        return None
    try:
        return datetime.strptime(元[:10], "%Y-%m-%d").date()
    except ValueError:
        pass
    # JavaScript の Date が文字列になった形（`Mon Mar 30 2026 00:00:00 GMT+0900 …`）
    m = re.match(r"^[A-Za-z]{3} ([A-Za-z]{3}) (\d{1,2}) (\d{4})", 元)
    if m:
        月 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].index(m.group(1)) + 1
        return datetime(int(m.group(3)), 月, int(m.group(2))).date()
    読めなかった日付.append(元[:40])
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
        self._取り込む(News, 生.get("news") or [], "お知らせ", 下見)
        self._取り込む(CalendarEvent, 生.get("calendar") or [], "カレンダー", 下見)
        self._取り込む(Menu, 生.get("menus") or [], "メニュー", 下見)
        self._取り込む(Product, 生.get("products") or [], "商品", 下見)
        self._カテゴリを取り込む(生.get("categories") or [], 下見)
        self._FAQを取り込む(生.get("faq") or [], 下見)
        self._プッシュ通知を取り込む(生.get("push_notices") or [], 下見)

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")

    def _取り込む(self, model, 行, 名, 下見):
        日付の鍵 = {News: "posted_on", CalendarEvent: "event_on",
                    Menu: "registered_on"}.get(model, "")
        # お知らせ・カレンダーは title、メニュー・商品は name。
        名前の鍵 = "title" if model in (News, CalendarEvent) else "name"
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0

        for r in 行:
            row = 数(r.get("row"))
            if not row:
                飛ばした += 1
                continue

            # 題も日付も無い行は、書きかけの空行。書き出し側でも弾いているが、
            # ここでも弾く。**受け取る側で守らないと、書き出しの作りが
            # 変わったときに空行が黙って入り込む。**
            if not 文字(r.get(名前の鍵)) and not (日付の鍵 and 日付(r.get(日付の鍵))):
                飛ばした += 1
                continue

            値 = {
                名前の鍵: 文字(r.get(名前の鍵))[:255],
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
            if model is Menu:
                値.update({
                    "registered_on": 日付(r.get("registered_on")),
                    "summary": 文字(r.get("summary")),
                    "category": 文字(r.get("category"))[:100],
                    "booking_status": 文字(r.get("booking_status"))[:100],
                    # 画像は配列のまま持つ。1つのメニューに複数枚ある。
                    "image_urls": r.get("image_urls") if isinstance(r.get("image_urls"), list) else [],
                    "sort_key": 数(r.get("sort_key")),
                })
            elif model is Product:
                値.update({
                    "category": 文字(r.get("category"))[:100],
                    "price": 数(r.get("price")),
                    "special_price": 数(r.get("special_price")),
                    "description": 文字(r.get("description")),
                    "icon_url": 文字(r.get("icon_url")),
                    "description_image_url": 文字(r.get("description_image_url")),
                    "background_color": 文字(r.get("background_color"))[:32],
                    # 空欄は None のまま。0（売り切れ）と区別する。
                    "stock": 数(r.get("stock")),
                    "stock_warning": 数(r.get("stock_warning")),
                    # 在庫数とは別物。院長が手で「売切」にできる。
                    # GAS はこの列だけを見て売切を判定している。
                    "sold_out": 文字(r.get("sold_out"))[:16],
                })
            elif model is News:
                値.update({
                    "posted_on": 日付(r.get("posted_on")),
                    "category": 文字(r.get("category"))[:100],
                    "icon": 文字(r.get("icon"))[:16],
                    "body": 文字(r.get("body")),
                    "link_url": 文字(r.get("link_url")),
                    "link_label": 文字(r.get("link_label"))[:100],
                    "button_text": 文字(r.get("button_text"))[:100],
                })
            elif model is CalendarEvent:
                値.update({
                    "event_on": 日付(r.get("event_on")),
                    "detail": 文字(r.get("detail")),
                    "color": 文字(r.get("color"))[:32],
                    "category": 文字(r.get("category"))[:100],
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

    def _FAQを取り込む(self, 行, 下見):
        """使い方FAQ。公開・非公開は「状態」列で、published という名前ではない。"""
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        # 数えるのは、実際に入る行だけ。飛ばした行まで数えた集計を出すと、
        # 「回答が空が2件」のように、直す必要のないものを直しに行かせてしまう。
        採用 = []

        for r in 行:
            row = 数(r.get("row"))
            質問 = 文字(r.get("question"))
            # 質問が無ければ答えようがない。GAS 側も question と answer が
            # 揃っていない行は使っていない（getSupportFaqEntries_）。
            if not row or not 質問:
                飛ばした += 1
                continue
            採用.append(r)

            値 = {
                "published": 公開か(r.get("status")),
                "category": 文字(r.get("category"))[:100],
                "question": 質問[:255],
                "keywords": 文字(r.get("keywords")),
                "answer": 文字(r.get("answer")),
                # 優先度が空の行は無かったが、空なら0として扱う。
                # 0なら最後に回るだけで、答えが消えるわけではない。
                "priority": 数(r.get("priority")) or 0,
                "updated_at": 日時(r.get("updated_at")),
            }

            既存 = SupportFaq.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ 使い方FAQ: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号か質問が空）: {飛ばした}件")

        公開数 = sum(1 for r in 採用 if 公開か(r.get("status")))
        回答なし = sum(1 for r in 採用 if not 文字(r.get("answer")))
        self.stdout.write(f"    うち公開: {公開数}件")
        if 回答なし:
            self.stdout.write(f"    **回答が空: {回答なし}件**（チャットが答えられません）")

        if 下見:
            return

        with transaction.atomic():
            for row, 値 in 新規:
                SupportFaq.objects.create(sheet_row=row, **値)
            for row, 値 in 更新:
                SupportFaq.objects.filter(sheet_row=row).update(**値)
        self.stdout.write(f"    → いま {SupportFaq.objects.count()}件")

    def _プッシュ通知を取り込む(self, 行, 下見):
        """送ったお知らせ通知の控え。

        **日時・通知ID・送信結果は、空欄を空欄のまま入れる。**
        埋めてしまうと「届いていない通知が届いたことになる」。
        """
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        採用 = []

        for r in 行:
            row = 数(r.get("row"))
            題 = 文字(r.get("title"))
            # 日時は96件中42件しか無いので、空行の手がかりにできない。
            # タイトルだけで判定する（書き出し側と同じ規則）。
            if not row or not 題:
                飛ばした += 1
                continue
            採用.append(r)

            値 = {
                "sent_at": 日時(r.get("sent_at")),
                "title": 題[:255],
                "body": 文字(r.get("body")),
                "target_status": 文字(r.get("target_status"))[:100],
                "target_detail": 文字(r.get("target_detail")),
                # 0件送信は実際にある。None と 0 を混ぜない。
                "recipient_count": 数(r.get("recipient_count")),
                "status": 文字(r.get("status"))[:32],
                "scheduled_at": 日時(r.get("scheduled_at")),
                "target_page": 文字(r.get("target_page"))[:32],
                "preview_body": 文字(r.get("preview_body")),
                "notification_id": 文字(r.get("notification_id"))[:64],
                "result": 文字(r.get("result")),
                "updated_at": 日時(r.get("updated_at")),
                "deleted": 削除済みか(r.get("deleted")),
                "deleted_at": 日時(r.get("deleted_at")),
            }

            既存 = PushNotice.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ プッシュ通知: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（行番号かタイトルが空）: {飛ばした}件")

        配信 = sum(1 for r in 採用 if 文字(r.get("notification_id")))
        日時あり = sum(1 for r in 採用 if 日時(r.get("sent_at")))
        self.stdout.write(f"    通知IDが返ってきた: {配信}件 / 日時がある: {日時あり}件")
        # この2つがずれていたら、書き出しか取り込みで取りこぼしている。
        if 配信 != 日時あり:
            self.stdout.write(
                f"    **通知IDと日時の件数が違います（{配信} と {日時あり}）。**"
                "シートでは揃っていたので、どこかで落ちています。"
            )

        if 下見:
            return

        with transaction.atomic():
            for row, 値 in 新規:
                PushNotice.objects.create(sheet_row=row, **値)
            for row, 値 in 更新:
                PushNotice.objects.filter(sheet_row=row).update(**値)
        self.stdout.write(f"    → いま {PushNotice.objects.count()}件")

    def _カテゴリを取り込む(self, 行, 下見):
        """カテゴリマスタ。2列しかないので、掲載の共通を持たない。"""
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0

        for r in 行:
            row = 数(r.get("row"))
            名 = 文字(r.get("name"))
            if not row or not 名:
                飛ばした += 1
                continue
            値 = {"name": 名[:100], "kind": 文字(r.get("kind"))[:32]}
            既存 = Category.objects.filter(sheet_row=row).first()
            if not 既存:
                新規.append((row, 値))
            elif [k for k, v in 値.items() if getattr(既存, k) != v]:
                更新.append((row, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ カテゴリ: JSONに {len(行)}件")
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした: {飛ばした}件")
        if 下見:
            return

        with transaction.atomic():
            for row, 値 in 新規:
                Category.objects.create(sheet_row=row, **値)
            for row, 値 in 更新:
                Category.objects.filter(sheet_row=row).update(**値)
        self.stdout.write(f"    → いま {Category.objects.count()}件")

        # **読めなかった日付を、必ず最後に報せる。**
        # 黙って None にしていたため、メニューの登録日13件が空のまま
        # 取り込みが「成功」していた（2026-09-05）。
        if 読めなかった日付:
            self.stdout.write("")
            self.stdout.write(self.style.WARNING(
                f"■ 読めなかった日付が {len(読めなかった日付)}件ありました。空のまま取り込んでいます。"))
            for x in 読めなかった日付[:5]:
                self.stdout.write(f"    {x}")
            if len(読めなかった日付) > 5:
                self.stdout.write(f"    …ほか {len(読めなかった日付) - 5}件")
