"""スプレッドシートから書き出した JSON を、会員台帳に取り込む。

    python manage.py 会員を取り込む 会員データ.json --下見
    python manage.py 会員を取り込む 会員データ.json

**会員IDで突き合わせるので、何度実行しても二重に増えない。**
すでにある方は、中身が変わっていれば更新する。

途中で止まっても、もう一度実行すれば揃う。「1回だけ実行してよい道具」は、
実行したかどうかを人が覚えていなければならず、いつか間違える。

パスコードは平文のまま持ち込まない。ここでハッシュにする。
すでにハッシュが入っている方は触らない（何度実行してもハッシュが
作り直されると、そのたびに値が変わって差分が出続けるため）。
"""

import json
import re as _re
from datetime import datetime

from django.contrib.auth.hashers import check_password, make_password
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from apps.members.models import Member


def 文字(値):
    return "" if 値 is None else str(値).strip()


def 数(値, 既定=0):
    try:
        return int(float(値))
    except (TypeError, ValueError):
        return 既定


def 真偽(値):
    s = 文字(値).lower()
    return s in ("true", "1", "on", "yes", "有効", "オン", "登録済み", "済")


def _電話(値):
    """電話番号。**消えた先頭の0を戻す。**

    スプレッドシートが電話番号を数値として受け取ると、先頭の0が落ちる。

        08012345678  →  8012345678

    2026-08-24 の時点で、153名のうち **121名**がこの形だった。
    表そのものは `gas/電話番号の0を戻す.js` で直したが、
    **ここでも受け止める。**片方だけに頼ると、また同じことが起きる。

    足すのは、足した結果が日本の電話番号の形になるときだけ。
    形にならないものは**そのまま持つ**（勝手に作り変えない）。
    """
    s = 文字(値)
    if not s:
        return ""
    数 = _re.sub(r"\D", "", s)
    if not 数 or 数.startswith("0"):
        return s

    # 国番号81（例: 817055600662 → 07055600662）
    if len(数) == 12 and 数.startswith("81"):
        return "0" + 数[2:]

    候補 = "0" + 数
    # 携帯（070/080/090 の11桁）
    if len(候補) == 11 and 候補[:3] in ("070", "080", "090"):
        return 候補
    # 固定電話（0で始まる10桁。07x は固定の市外局番に無い）
    if len(候補) == 10 and 候補[1] != "0" and not 候補.startswith("07"):
        return 候補

    # 形にならないものは、そのまま。**推測で作り変えない。**
    return s


def _届け先(値):
    """通知の届け先。真偽値の false と文字列の "false" は、どちらも空にする。"""
    s = 文字(値)
    return "" if s.lower() in ("false", "0", "none", "null", "undefined") else s


def _通知が入か(値):
    """届け先の中身から、通知オンかどうかを決める。

    GAS は `!!row[USER_COL.PUSH - 1]`（7933行）。セルから読むと真偽値の
    false がそのまま来るので偽になる。**こちらは文字で受け取る**ので、
    "false" という文字列を明示的に偽へ寄せる。
    """
    return bool(_届け先(値))


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


def 一覧(値):
    """JSON文字列でも配列でも受け取る。壊れていても止めない。

    履歴のJSONは長年ぶんが入っており、一部が壊れていることがある。
    そこで止めると、その方だけ移せずに残ってしまう。
    読めなければ空にして、あとで数を突き合わせたときに気づけるようにする。
    """
    if isinstance(値, list):
        return 値
    s = 文字(値)
    if not s:
        return []
    try:
        v = json.loads(s)
        return v if isinstance(v, list) else []
    except (ValueError, TypeError):
        return []


class Command(BaseCommand):
    help = "会員データのJSONを会員台帳に取り込む"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview",
                            help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        path = options["json_path"]
        下見 = options["preview"]

        try:
            with open(path, encoding="utf-8") as f:
                生 = json.load(f)
        except OSError as e:
            raise CommandError(f"読み込めませんでした: {e}")
        except ValueError as e:
            raise CommandError(f"JSONとして読めませんでした: {e}")

        行 = 生.get("members") if isinstance(生, dict) else 生
        if not isinstance(行, list):
            raise CommandError("members の配列が見つかりません。")

        新規, 更新, 変化なし, 飛ばした = [], [], 0, []
        パスコードを作った = 0

        for r in 行:
            member_id = 文字(r.get("memberId") or r.get("ID") or r.get("id"))
            name = 文字(r.get("name") or r.get("氏名"))
            if not member_id:
                飛ばした.append(f"会員IDが空: {name or '（お名前も空）'}")
                continue
            if not name:
                飛ばした.append(f"お名前が空: {member_id}")
                continue

            既存 = Member.objects.filter(pk=member_id).first()

            値 = {
                "name": name,
                "kana": 文字(r.get("kana")),
                "phone": _電話(r.get("phone")),
                "birthday": 日付(r.get("birthday")),
                "address": 文字(r.get("address")),
                "avatar_url": 文字(r.get("avatarUrl")),
                # **8列目は「オン/オフ」ではなく、通知の届け先そのもの。**
                # 2026-08-23 まで、ここで 真偽() に通していた。
                # 購読IDは "true" でも "1" でもないので**偽に落ちていた。**
                "push_subscription": _届け先(r.get("pushSubscription")
                                             if "pushSubscription" in r
                                             else r.get("pushEnabled")),
                "push_enabled": _通知が入か(r.get("pushSubscription")
                                            if "pushSubscription" in r
                                            else r.get("pushEnabled")),
                "memo": 文字(r.get("memo")),
                "stamp_achieved_at": 日時(r.get("stampAchievedAt")),
                "status": 文字(r.get("status")),
                "stamp_count": 数(r.get("stampCount")),
                "stamp_card_number": 数(r.get("stampCardNumber")),
                "last_stamp_at": 日時(r.get("lastStampAt")),
                "password_hash": 文字(r.get("passwordHash")),
                "password_salt": 文字(r.get("passwordSalt")),
                "role": 文字(r.get("role")),
                "transfer_code": 文字(r.get("transferCode")),
                "transfer_code_issued_at": 日時(r.get("transferCodeIssuedAt")),
                "device_sessions": 一覧(r.get("deviceSessions")),
                "stamp_history": 一覧(r.get("stampHistory")),
                "reward_history": 一覧(r.get("rewardHistory")),
                "deleted": 真偽(r.get("deleted")),
                "deleted_at": 日時(r.get("deletedAt")),
                "merged_into_id": 文字(r.get("mergedIntoId")),
                "registration_source": 文字(r.get("registrationSource")),
                "registration_source_detail": 文字(r.get("registrationSourceDetail")),
                "registration_source_updated_at": 日時(r.get("registrationSourceUpdatedAt")),
                "last_online_at": 日時(r.get("lastOnlineAt")),
                # 空文字ではなく None にする。unique 制約は空文字を重複と見なすため、
                # **2人目で必ず落ちる。**
                "line_user_id": (文字(r.get("lineUserId")) or None),
                "bijiris_registered": 真偽(r.get("bijirisRegistered")),
                "created_at": 日時(r.get("createdAt")),
            }

            # 平文のパスコードは持ち込まない。ここでハッシュにする。
            # すでに入っている方は触らない（毎回作り直すと値が変わり続けるため）。
            # パスコードは平文で持ち込まない。ここでハッシュにする。
            #
            # **すでにハッシュがあっても、合わなければ入れ直す。**
            # 2026-08-27、表のパスコード36件で**先頭の0が消えていた**ことが分かった
            # （0123 が 123 になっていた）。表を直しても、ここで
            # 「すでにハッシュがあるから」と飛ばすと、**古い間違ったハッシュが残る。**
            # GAS は照合時に0を補うので通っていたが、**ハッシュは補えない。**
            # 切り替えた瞬間に、その36名が入れなくなる。
            平文 = 文字(r.get("passcode"))
            if 平文:
                古い = 既存.passcode_hash if 既存 else ""
                if not 古い or not check_password(平文, 古い):
                    値["passcode_hash"] = make_password(平文)
                パスコードを作った += 1

            if not 既存:
                新規.append((member_id, 値, name))
                continue

            変わった = [
                k for k, v in 値.items()
                if k != "passcode_hash" and getattr(既存, k) != v
            ]
            if 変わった:
                更新.append((member_id, 値, name, 変わった))
            else:
                変化なし += 1

        # ---- 見せる ----
        self.stdout.write("")
        self.stdout.write(f"■ JSONの中の会員: {len(行)}名")
        self.stdout.write(f"    新しく入る:   {len(新規)}名")
        self.stdout.write(f"    中身が変わる: {len(更新)}名")
        self.stdout.write(f"    変わらない:   {変化なし}名")
        if パスコードを作った:
            self.stdout.write(f"    パスコードをハッシュにする: {パスコードを作った}名")
        if 飛ばした:
            self.stdout.write("")
            self.stdout.write(f"■ 取り込めない行: {len(飛ばした)}件")
            for x in 飛ばした[:20]:
                self.stdout.write(f"    {x}")

        for member_id, _, name, 変わった in 更新[:20]:
            self.stdout.write(f"    変更 {member_id} {name}: {'・'.join(変わった[:6])}")
        if len(更新) > 20:
            self.stdout.write(f"    …ほか {len(更新) - 20}名")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。")
            self.stdout.write("  よければ --下見 を外して実行してください。")
            return

        # ---- 書く ----
        with transaction.atomic():
            for member_id, 値, _ in 新規:
                Member.objects.create(member_id=member_id, **値)
            for member_id, 値, _, _ in 更新:
                Member.objects.filter(pk=member_id).update(**値)

        self.stdout.write("")
        self.stdout.write(f"■ 取り込みました。会員台帳はいま {Member.objects.count()}名です。")
        self.stdout.write("  スプレッドシートの件数と合っているか確かめてください。")
