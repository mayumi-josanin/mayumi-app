"""ビジリスの設定（ADMIN_PREFERENCES_JSON / TICKET_SURVEY_META_JSON / TICKET_SURVEY_PROMPT）を取り込む。

    python manage.py ビジリスの設定を取り込む ビジリス設定.json --下見
    python manage.py ビジリスの設定を取り込む ビジリス設定.json

JSON は gas/ビジリスを書き出す.js の `ビジリスの設定を書き出す()`（ビジリスの GAS で動かす）。
形は {"preferences": {…}, "ticket_meta": {…}, "ticket_prompt": "…"}。

records.AppSetting に3つの鍵で入れる:
  bijiris_preferences  … getPreferences_ と同じ既定で欠けを埋めたもの（apps/bijiris/preferences.py）
  bijiris_ticket_meta  … 回数券分析の進行状況（中身そのまま）
  bijiris_ticket_prompt … {"prompt": "…"}（空なら既定の指示文を使う。空のまま持つ）

**秘密は入れない。**ANTHROPIC_API_KEY / ONESIGNAL_* / TOKEN_SECRET / ADMIN_PASSWORD などが JSON に混ざっていても
捨てて数える（書き出す側でも出さない）。控えに秘密が混ざるため。
"""

from django.core.management.base import BaseCommand
from django.db import transaction

from apps.bijiris import preferences
from apps.records.models import AppSetting

from ._common import ファイルを開く, 文字

秘密の鍵 = ("ANTHROPIC_API_KEY", "ONESIGNAL_REST_API_KEY", "ONESIGNAL_APP_ID", "TOKEN_SECRET", "ADMIN_PASSWORD",
        "ADMIN_USERNAME", "ADMIN_USERS_JSON", "apiKey", "restApiKey")


def _置く(鍵, 値, 覚え書き, 下見, out):
    既存 = AppSetting.objects.filter(pk=鍵).first()
    if 既存 and 既存.value == 値:
        out.write(f"    {鍵}: 変わらない")
        return
    out.write(f"    {鍵}: {'新しく入る' if not 既存 else '中身が変わる'}")
    if not 下見:
        AppSetting.objects.update_or_create(pk=鍵, defaults={"value": 値, "note": 覚え書き})


class Command(BaseCommand):
    help = "ビジリスの設定の JSON を取り込む（AppSetting bijiris_preferences / bijiris_ticket_meta / bijiris_ticket_prompt）"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview", help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        生 = ファイルを開く(options["json_path"])
        if not isinstance(生, dict):
            生 = {}
        下見 = options["preview"]

        混ざった秘密 = [k for k in 秘密の鍵 if k in 生 or k in (生.get("preferences") or {})]
        pref生 = 生.get("preferences") if isinstance(生.get("preferences"), dict) else {}
        pref生 = {k: v for k, v in pref生.items() if k not in 秘密の鍵}
        pref = preferences.正規化する(pref生)
        meta = 生.get("ticket_meta") if isinstance(生.get("ticket_meta"), dict) else {}
        prompt = 文字(生.get("ticket_prompt"))

        self.stdout.write("")
        self.stdout.write("■ ビジリスの設定")
        self.stdout.write(f"    通知: {'出す' if pref['notificationEnabled'] else '出さない'} → {pref['notificationEmail'] or '（空）'}")
        self.stdout.write(f"    特典（節目）: {'有効' if pref['milestoneRewardConfig']['enabled'] else '無効'}"
                          f" {len(pref['milestoneRewardConfig']['milestones'])}段 / ガチャ {len(pref['gachaPrizeConfig']['monthlyPrizes'])}か月ぶん"
                          f" / キャンペーンスタンプ {'有効' if pref['campaignStampEnabled'] else '無効'}")
        self.stdout.write(f"    回数券分析のメタ: {len(meta)}項目 / 指示文: {'あり ' + str(len(prompt)) + '文字' if prompt else '空（既定を使う）'}")
        if 混ざった秘密:
            self.stdout.write(f"    **秘密が混ざっていたので捨てました: {'・'.join(混ざった秘密)}**（.env に置く）")
        with transaction.atomic():
            _置く(preferences.鍵, pref, "ビジリスの設定（ADMIN_PREFERENCES_JSON の写し）", 下見, self.stdout)
            _置く(preferences.回数券メタの鍵, meta, "ビジリス 回数券分析の進行状況（TICKET_SURVEY_META_JSON の写し）", 下見, self.stdout)
            _置く(preferences.回数券指示文の鍵, {"prompt": prompt}, "ビジリス 回数券分析の AI への指示文（TICKET_SURVEY_PROMPT の写し）", 下見, self.stdout)

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。よければ --下見 を外して実行してください。")
