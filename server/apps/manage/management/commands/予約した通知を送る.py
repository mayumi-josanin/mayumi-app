"""予約したお知らせ通知を、時間が来たらサーバーから送る。

    python manage.py 予約した通知を送る          … 時間が来た予約を送る
    python manage.py 予約した通知を送る --下見    … 送らずに、何件・どれを送るかだけ出す

## なぜサーバーから送るか（2026-09-20）

お知らせの「予約」は管理画面で記録されるが、**時間が来たときに送る役目は
GAS の自動実行（processScheduledPushQueue）だけ**が担っていた。
その GAS が 9/18〜20 の2日間で6回、名前引き（DNS）の失敗で落ちていた。
予約が無かったので実害は出ていないが、次の予約は時間どおりに届かない恐れがある。
**サーバーが自分で送れるようにして、GAS の調子に左右されないようにする。**
（docker-compose.yml の cron が scripts/scheduler.sh 経由で1分おきに呼ぶ）

## 二度送らないための決めごと

送る前に **先に「送った」印（sent_at と状態）を立てて保存する。**
GAS の自動実行は同じ表をサーバー越しに見ていて、**状態が「予約済み」の行だけ**を拾う。
先に印を立てておけば、向こうは同じ行を拾わない。
行を掴む（select_for_update）のと合わせて、同じ分に二重に走っても片方しか送れない。

送信に失敗したら印（sent_at）を戻し、「送信失敗」と理由を残して次の行へ進む。
1件の失敗で残り全部を止めない。
**「予約済み」には戻さない。**戻すと1分後にまた送ろうとして、
届かないまま毎分やり直し続ける（届いたのに答えだけ落ちた場合は二度届く）。
失敗は履歴に残るので、院長が同じ内容で送り直せる。
"""

from django.core.management.base import BaseCommand
from django.db import transaction
from django.utils import timezone

from apps.content import onesignal
from apps.content.models import PushNotice
from apps.gasapi import admin_push
from apps.manage import views_push

鍵が無い = "通知の鍵が設定されていないため送れませんでした（.env の ONESIGNAL_APP_ID / ONESIGNAL_REST_API_KEY）"


def 時間が来た予約():
    """予約日時が過ぎていて、まだ送っていないもの。古い順。

    **予約日時が空のものは触らない。**GAS の processScheduledPushQueue は
    空を「もう来ている」扱いにして送ってしまうが、日時を入れ忘れた行が
    突然届くことになる。下書き・送信済み・削除済みも状態で外れる。
    """
    return (
        PushNotice.objects
        .filter(status=admin_push.SCHEDULED, deleted=False,
                sent_at__isnull=True, scheduled_at__isnull=False,
                scheduled_at__lte=timezone.now())
        .order_by("scheduled_at", "sheet_row")
    )


def 届け先の端末(記録: PushNotice, 会員一覧) -> list:
    """送り先の絞り込み。**管理画面（views_push）と同じ考え方。**

    全員のときは端末を指定しない（OneSignal のタグで全員へ）。
    絞ったときだけ、いまの会員から絞り直して端末を並べる。
    予約したときの一覧をそのまま使わないのは、予約から実際に送るまでの間に
    注文や特典の状況が変わるため。画面で送るときと同じ結果にする。
    """
    mode = (記録.target_status or "all").strip() or "all"
    if mode == "all":
        return []
    return [str(u.get("subscription") or "").strip()
            for u in views_push._絞る(会員一覧, mode)
            if str(u.get("subscription") or "").strip()]


def 一件送る(記録: PushNotice, 会員一覧) -> tuple[bool, str]:
    """印を先に立ててから送る。送れなければ印を戻して「送信失敗」にする。"""
    ids = 届け先の端末(記録, 会員一覧)
    try:
        答 = onesignal.送信(onesignal.内容(記録.title, 記録.body, 記録.target_page, ids))
    except onesignal.送信できない as e:
        記録.sent_at = None
        記録.status = admin_push.FAILED
        記録.result = str(e)[:1000]
        記録.save(update_fields=["sent_at", "status", "result", "changed_at"])
        return False, str(e)
    記録.notification_id = admin_push._文(答.get("id"))[:64]
    受取 = int(答.get("recipients") or 0)
    if 受取:
        記録.recipient_count = 受取
    記録.result = "予約送信済み"
    記録.save(update_fields=["notification_id", "recipient_count", "result", "changed_at"])
    return True, "予約送信済み"


class Command(BaseCommand):
    help = "予約したお知らせ通知のうち、時間が来たものを送る"

    def add_arguments(self, parser):
        parser.add_argument("--下見", action="store_true",
                            help="送らずに、何件・どれを送るかだけ出す")

    def handle(self, *args, **opts):
        下見 = bool(opts.get("下見"))
        待ち = list(時間が来た予約())
        if not 待ち:
            if 下見:
                self.stdout.write("送る予約はありません")
            return

        if 下見:
            self.stdout.write(f"送る予約 {len(待ち)}件")
            for p in 待ち:
                予定 = timezone.localtime(p.scheduled_at).strftime("%Y/%m/%d %H:%M")
                self.stdout.write(f"  {p.sheet_row}行 {予定} {p.title}（{p.target_status or 'all'}）")
            return

        if not onesignal.設定済みか():
            # 鍵が無いのに送ろうとしても必ず失敗する。**落とさずに理由を残して終わる。**
            # 状態は「予約済み」のまま。鍵を入れれば次の1分で送られる。
            for p in 待ち:
                p.result = 鍵が無い
                p.save(update_fields=["result", "changed_at"])
            self.stdout.write(鍵が無い + f"（{len(待ち)}件そのまま）")
            return

        # 届け先の一覧は1回だけ取る（1分おきに走るので、毎行取り直さない）
        会員一覧 = views_push._届け先()
        送れた = 送れなかった = 0
        for p in 待ち:
            with transaction.atomic():
                記録 = (PushNotice.objects.select_for_update(skip_locked=True)
                        .filter(pk=p.pk, status=admin_push.SCHEDULED, sent_at__isnull=True,
                                deleted=False).first())
                if not 記録:
                    continue  # ほかの実行が先に掴んだ
                # **ここが二度送らない要。**送る前に印を立てて確定させる。
                記録.sent_at = timezone.now()
                記録.status = admin_push.SENT
                記録.save(update_fields=["sent_at", "status", "changed_at"])
            ok, 結果 = 一件送る(記録, 会員一覧)
            if ok:
                送れた += 1
                self.stdout.write(f"送りました: {記録.sheet_row}行 {記録.title}")
            else:
                送れなかった += 1
                self.stderr.write(f"送れませんでした: {記録.sheet_row}行 {記録.title} — {結果}")
        self.stdout.write(f"予約した通知: {送れた}件送信 / {送れなかった}件失敗")
