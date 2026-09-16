"""ビジリスの顧客プロフィール（CUSTOMER_PROFILES_JSON）とメモ（CUSTOMER_MEMOS_JSON）を取り込む。

    python manage.py ビジリスの顧客を取り込む ビジリス顧客.json --下見
    python manage.py ビジリスの顧客を取り込む ビジリス顧客.json

JSON は gas/ビジリスを書き出す.js の `ビジリスの顧客プロフィールを書き出す()`（ビジリスの GAS で動かす）。
形は {"profiles": {お名前: {…}}, "memos": {お名前: {latestMemo, entries} | "文字列"}}（プロパティの生のまま）。

**お名前で突き合わせる**（GAS の鍵がお名前なので、ここだけはそう）。何度実行しても二重に増えない。
**お名前で会員に結びつけない。**member_id は空のまま。GAS が持っている memberNumber は
`member_number` に写すだけ（GAS がまゆみの会員データを氏名で引いた結果で、確かめていない）。

GAS の `normalizeCustomerProfileRecord_` は「お名前になっていない」（かな・漢字・英数字を1文字も含まない）
鍵を捨てる。同じにする。
"""

import re

from django.core.management.base import BaseCommand
from django.db import transaction

from apps.bijiris.models import CustomerProfile
from apps.members.models import Member

from ._common import ファイルを開く, 文字, 日時, 違い

お名前の文字 = re.compile(r"[ぁ-んァ-ヶ一-龥a-zA-Z0-9]")


def お名前になっていない(名: str) -> bool:
    return not 名 or not お名前の文字.search(名)


def 一覧(値) -> list:
    return list(dict.fromkeys(v for v in (文字(x) for x in (値 if isinstance(値, list) else [])) if v))


def メモを整える(値) -> dict:
    """normalizeCustomerMemoRecord_ と同じ。文字列なら1件のメモ。新しい順。"""
    if not 値:
        return {"latestMemo": "", "entries": []}
    if isinstance(値, str):
        m = 文字(値)
        return {"latestMemo": m, "entries": [{"at": "", "memo": m}] if m else []}
    if not isinstance(値, dict):
        return {"latestMemo": "", "entries": []}
    latest = 文字(値.get("latestMemo") or 値.get("memo"))
    entries = []
    for e in (値.get("entries") if isinstance(値.get("entries"), list) else []):
        if not isinstance(e, dict):
            continue
        memo = 文字(e.get("memo") or e.get("content"))
        if memo:
            entries.append({"at": 文字(e.get("at"))[:10], "memo": memo})
    entries.sort(key=lambda x: x["at"], reverse=True)
    if entries:
        latest = entries[0]["memo"]
    return {"latestMemo": latest, "entries": entries[:100]}


def 行にする(名: str, p: dict, メモ) -> dict:
    出どころ = 文字(p.get("activeTicketCardSource") or p.get("ticketCardSource")).lower()
    m = メモを整える(メモ)
    adj = p.get("ticketStampAdjustment")
    try:
        adj = max(-50, min(50, int(float(adj))))
    except (TypeError, ValueError):
        adj = 0
    return {
        "member_id": "",
        "member_number": 文字(p.get("memberNumber"))[:32],
        "name_kana": re.sub(r"\s+", "", 文字(p.get("nameKana")))[:255],
        "aliases": [a for a in 一覧(p.get("aliases")) if a != 名 and not お名前になっていない(a)],
        "client_ids": 一覧(p.get("clientIds")),
        "active_ticket_card": p.get("activeTicketCard") if isinstance(p.get("activeTicketCard"), dict) else None,
        "active_ticket_card_source": 出どころ if 出どころ in ("admin", "response", "customer") else "",
        "last_ticket_card_acquired_at": 日時(p.get("lastTicketCardAcquiredAt") or p.get("ticketCardLastAcquiredAt")),
        "measurement_targets": p.get("measurementTargets") if isinstance(p.get("measurementTargets"), dict) else None,
        "ticket_stamp_adjustment": adj,
        "push_status": p.get("pushStatus") if isinstance(p.get("pushStatus"), dict) else None,
        "reward_redemptions": p.get("rewardRedemptions") if isinstance(p.get("rewardRedemptions"), dict) else None,
        "admin_managed": p.get("adminManaged") is True,
        "passcode_hash": 文字(p.get("passcodeHash"))[:128],
        "passcode_salt": 文字(p.get("passcodeSalt"))[:128],
        "passcode_updated_at": 文字(p.get("passcodeUpdatedAt"))[:64],
        "passcode_setup_until": 文字(p.get("passcodeSetupUntil"))[:64],
        "latest_memo": m["latestMemo"],
        "memo_entries": m["entries"],
        "updated_at": 日時(p.get("updatedAt")),
    }


class Command(BaseCommand):
    help = "ビジリスの顧客プロフィールとメモの JSON を取り込む"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview", help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        生 = ファイルを開く(options["json_path"])
        profiles = 生.get("profiles") if isinstance(生, dict) and isinstance(生.get("profiles"), dict) else {}
        memos = 生.get("memos") if isinstance(生, dict) and isinstance(生.get("memos"), dict) else {}
        下見 = options["preview"]

        新規, 更新, 変化なし, 捨てた = [], [], 0, []
        名前ら = set(profiles) | set(memos)
        for 鍵 in sorted(名前ら):
            p = profiles.get(鍵) if isinstance(profiles.get(鍵), dict) else {}
            名 = 文字(p.get("name")) or 文字(鍵)
            if お名前になっていない(名):
                捨てた.append(鍵)
                continue
            値 = 行にする(名, p, memos.get(鍵) if 鍵 in memos else memos.get(名))
            既存 = CustomerProfile.objects.filter(name=名[:255]).first()
            if not 既存:
                新規.append((名[:255], 値))
            elif 違い(既存, 値):
                更新.append((名[:255], 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ ビジリスの顧客: プロフィール {len(profiles)}名 / メモ {len(memos)}名（お名前の種類 {len(名前ら)}）")
        self.stdout.write(f"    新しく入る:   {len(新規)}名")
        self.stdout.write(f"    中身が変わる: {len(更新)}名")
        self.stdout.write(f"    変わらない:   {変化なし}名")
        if 捨てた:
            self.stdout.write(f"    お名前になっていない鍵（GAS も捨てる）: {len(捨てた)}件 {捨てた[:5]}")
        全部 = 新規 + 更新
        self.stdout.write(f"    パスコード設定済み: {sum(1 for _, v in 全部 if v['passcode_hash'] and v['passcode_salt'])}名"
                          f" / 回数券カードあり: {sum(1 for _, v in 全部 if v['active_ticket_card'])}名"
                          f" / 特典の受け取りあり: {sum(1 for _, v in 全部 if v['reward_redemptions'])}名"
                          f" / メモあり: {sum(1 for _, v in 全部 if v['latest_memo'])}名"
                          f" / GAS の会員番号あり: {sum(1 for _, v in 全部 if v['member_number'])}名")

        # **お名前の照合は数えるだけ。結びつけない。**
        一致 = sum(1 for n in 名前ら if not お名前になっていない(n) and Member.objects.filter(name=n).count() == 1)
        self.stdout.write(f"  ● 会員のお名前との一致（**数えるだけ。結びつけていません**）: ちょうど1名と一致 {一致}名")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for 名, 値 in 新規:
                CustomerProfile.objects.create(name=名, **値)
            for 名, 値 in 更新:
                CustomerProfile.objects.filter(name=名).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {CustomerProfile.objects.count()}名")
