"""お客様の入口。**第3段。いちばん危ない。**

計画は [会員の表を移す](docs/design/会員の表を移す.md)。

    loginAccount        ログイン
    registerAccount     新規登録
    registerBijirisUse  ビジリスの利用登録

## ここが止まると、お客様がアプリに入れなくなる

だから、第1段（読むだけ）と第2段（管理側の書き込み）が
**数日おかしくないことを見てから**切り替える。

## GAS との違いを、はっきりさせておく

| | GAS | サーバー |
|---|---|---|
| パスコード | **平文で保存**し、平文で比べる | **ハッシュ**（`check_password`） |
| 壊れた0の修復 | 照合が通ったら書き直す | **要らない**（ハッシュなので壊れない） |
| 回数制限 | スクリプトプロパティ | データベース（処理が並んでも効く） |

**平文をやめるのが、この移行の目的の1つ。**

## 会員IDは `MYM-` ＋ 4桁

GAS は `MYM-` + 1000〜9999 の乱数で、使われていないものを探す。
**9000通りしかない。**いま164名なので当分は足りるが、
**採番に失敗したら断る**（黙って番号を作らない）。

**会員IDは変えない。**変えるとお客様が全員入り直しになる。
"""

import base64
import hashlib
import hmac
import json
import random
import re
import time
import unicodedata

from django.conf import settings
from django.contrib.auth.hashers import check_password, make_password
from django.core.cache import cache
from django.db import transaction
from django.utils import timezone

from apps.members.models import Member

ログインの上限 = 10
ログインのお休み秒 = 600


def _文(v):
    return "" if v is None else str(v)


def _名前で照らす(値):
    """比べるためだけの形。**空白と全角半角の違いを消す。**

    GAS の `normalizeNameForMatch_` と同じにすること。
    ずれると、ご本人が「お名前が違います」で入れなくなる。
    """
    return re.sub(r"[\s　]+", "", unicodedata.normalize("NFKC", _文(値)))


def _日付の文字(値):
    """`YYYY-MM-DD` にそろえる。GAS の normalizeDateOnlyValue_ と同じ。"""
    文 = _文(値).strip()
    if not 文:
        return ""
    文 = 文.replace("/", "-")[:10]
    return 文 if re.match(r"^\d{4}-\d{2}-\d{2}$", 文) else ""


def _パスコードとして正しいか(値):
    """**数字4桁か6桁。**GAS の isValidPasscodeValue_ と同じ。"""
    return bool(re.match(r"^\d{4}$|^\d{6}$", _文(値).strip()))


def _電話を整える(文):
    """消えた先頭の0を戻す。"""
    値 = _文(文).strip()
    数 = re.sub(r"\D", "", 値)
    if not 数 or 数.startswith("0"):
        return 値
    if len(数) == 12 and 数.startswith("81"):
        return "0" + 数[2:]
    候補 = "0" + 数
    if len(候補) == 11 and 候補[:3] in ("070", "080", "090"):
        return 候補
    if len(候補) == 10 and 候補[1] != "0" and not 候補.startswith("07"):
        return 候補
    return 値


# ── 回数制限（お名前ごと。全体では止めない） ─────────────

def _鍵(名):
    """**お名前を平文で残さない。**記録そのものが漏れうる情報になる。"""
    return "login_" + hashlib.sha256(_文(名).encode("utf-8")).hexdigest()[:24]


def _お休み中か(名):
    try:
        return int(cache.get(_鍵(名)) or 0) >= ログインの上限
    except Exception:
        # **数えられないときにログインを止めない。**
        # 止めると、本当のご本人が締め出される。
        return False


def _はずれを数える(名):
    try:
        k = _鍵(名)
        cache.set(k, int(cache.get(k) or 0) + 1, ログインのお休み秒)
    except Exception:
        pass


def _数えるのをやめる(名):
    try:
        cache.delete(_鍵(名))
    except Exception:
        pass


# ── ログイン ─────────────────────────────────────────

# ── 入口が返す答え（GAS の buildAccountSession_ と同じ形）────────
#
# **札の署名はサーバーで行う。**GAS と同じ秘密（ADMIN_TOKEN_SECRET）・同じ作り
# （HMAC-SHA256 → base64url）なので、GAS の verifyMemberToken_ / verifyAdminToken_ が
# そのまま本物と認める。札の形を変えるとお客様が全員入り直しになる（CLAUDE.md 9）。
# 2026-09-15 の試験で「サーバーは memberId/name/kana しか返さず、入口は token が
# 無いとログイン失敗にする」ことが分かった（start/index.html 677行）。

札の寿命ミリ秒 = 90 * 24 * 60 * 60 * 1000  # GAS の ACCOUNT_TOKEN_TTL_MS と同じ


class 札の秘密がない(Exception):
    pass


def _秘密():
    秘密 = getattr(settings, "ADMIN_TOKEN_SECRET", "") or ""
    if not 秘密:
        raise 札の秘密がない()
    return 秘密


def _base64url(生: bytes) -> str:
    # GAS の Utilities.base64EncodeWebSafe は詰め物（=）を付ける。Python も同じ。
    return base64.urlsafe_b64encode(生).decode("ascii")


def _署名(値: str) -> str:
    return _base64url(hmac.new(_秘密().encode("utf-8"), 値.encode("utf-8"), hashlib.sha256).digest())


def _札にする(中身: dict) -> str:
    return _base64url(json.dumps(中身, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def _会員の札(会員ID: str, 期限ミリ秒: int) -> str:
    return _札にする({"t": "member", "u": 会員ID, "exp": 期限ミリ秒,
                     "sig": _署名(f"member|{会員ID}|{期限ミリ秒}")})


def _管理者の札(期限ミリ秒: int) -> str:
    return _札にする({"t": "admin", "exp": 期限ミリ秒, "sig": _署名(f"admin|{期限ミリ秒}")})


def _漢字が要るか(名: str) -> bool:
    """GAS の needsKanjiName_。ひらがな（と長音）だけのお名前は、あとから漢字で
    登録し直されて二重になりやすいので入口で止める。"""
    s = re.sub(r"[\s　]", "", _文(名))
    if not s:
        return False
    if re.search(r"[\u4E00-\u9FFF\u3005]", s):
        return False
    return bool(re.fullmatch(r"[\u3041-\u309F\u30FC]+", s))


def _保存する名前(値) -> str:
    """GAS の normalizeStoredName_。空白だけ取る。"""
    return re.sub(r"[\s　]+", "", _文(値))


def _保存するかな(値) -> str:
    """GAS の normalizeStoredKana_。NFKC → 空白を取る → カタカナをひらがなへ。"""
    文 = re.sub(r"[\s　]+", "", unicodedata.normalize("NFKC", _文(値)))
    return "".join(chr(ord(c) - 0x60) if "\u30A1" <= c <= "\u30F6" else c for c in 文)


def _足りない項目(m) -> list:
    """GAS の missingProfileFields_。値そのものは返さない。入口がお願いする材料。"""
    足りない = []
    if _漢字が要るか(m.name):
        足りない.append("name")
    if not (m.kana or "").strip():
        足りない.append("kana")
    if not (m.phone or "").strip():
        足りない.append("phone")
    if not m.birthday:
        足りない.append("birthday")
    if not (m.address or "").strip():
        足りない.append("address")
    return 足りない


def _入口の答え(m, **追加) -> dict:
    """GAS の buildAccountSession_ と同じ項目。"""
    期限 = int(time.time() * 1000) + 札の寿命ミリ秒
    管理者 = (m.role or "").strip() == "管理者"
    try:
        会員札 = _会員の札(m.member_id, 期限)
        札 = _管理者の札(期限) if 管理者 else 会員札
    except 札の秘密がない:
        # 既定値へ落とさない（GAS と同じ。安全側）。
        return {"status": "error",
                "message": "ログインの準備ができていません（サーバーの設定）。受付にお申し出ください。"}
    秒 = 期限 // 1000
    答 = {
        "status": "ok",
        "role": "admin" if 管理者 else "member",
        "memberId": m.member_id,
        "name": m.name or "",
        "kana": m.kana or "",
        "bijiris": bool(m.bijiris_registered),
        "token": 札,
        "memberToken": 会員札,
        # JavaScript の toISOString と同じ書き方（ミリ秒つき・Z）。
        "expiresAt": time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(秒)) + ".%03dZ" % (期限 % 1000),
        "missing": _足りない項目(m),
    }
    答.update(追加)
    return 答


def _パスコードが合う(m, パス: str) -> bool:
    """いまの形（ハッシュ）か、旧いパスワードの形（GAS の hashAccountPassword_）。"""
    if m.passcode_hash and check_password(パス, m.passcode_hash):
        return True
    if m.password_hash and m.password_salt:
        # GAS 2207行: sha256(salt|password|ADMIN_TOKEN_SECRET) の base64url。
        # 2026-09-15 まで sha256(password+salt) の hex で比べており、旧方式の2名が入れなかった。
        try:
            素 = (m.password_salt + "|" + パス + "|" + _秘密()).encode("utf-8")
            if _base64url(hashlib.sha256(素).digest()) == m.password_hash:
                return True
        except 札の秘密がない:
            pass
        # 以前のサーバーの作り。取り込んだ値がこの形で残っている可能性に備えて残す。
        種 = (パス + m.password_salt).encode("utf-8")
        if hashlib.sha256(種).hexdigest() == m.password_hash:
            return True
    return False


def ログイン(d):
    """GAS の loginAccount。**お名前とパスコードで探す。**"""
    生の名 = _文(d.get("name")).strip()
    名 = _名前で照らす(生の名)
    パス = _文(d.get("passcode") or d.get("password")).strip()
    if not 名 or not パス:
        return {"status": "error", "message": "お名前とパスコードを入力してください。"}

    if _お休み中か(名):
        return {"status": "error",
                "message": "ログイン失敗が続いたため、一時的に停止しています。10分後に再試行してください。"}

    当たり = [m for m in Member.objects.exclude(deleted=True)
            if _名前で照らす(m.name) == 名 and _パスコードが合う(m, パス)]

    # **ちょうど1人でなければ通さない。**
    # 0人は「違う」、2人以上は同姓同名で見分けがつかない。
    if len(当たり) != 1:
        _はずれを数える(名)
        return {"status": "error", "message": "お名前またはパスコードが違います。"}

    m = 当たり[0]
    _数えるのをやめる(名)
    m.last_online_at = timezone.now()
    m.save(update_fields=["last_online_at", "changed_at"])
    return _入口の答え(m)


# ── 新規登録 ─────────────────────────────────────────

def 新規登録(d):
    """GAS の registerAccount。

    **すでにある記録を引き受ける場合がある。**受付が先に作った行に、
    お客様があとからパスコードを付ける形。お名前と生年月日で照らす。
    守り（GAS 2378〜2450行）も同じ順で持つ。
    """
    生の名 = _文(d.get("name")).strip()
    名 = _名前で照らす(生の名)
    かな = _文(d.get("kana")).strip()
    生年月日 = _日付の文字(d.get("birthday"))
    電話 = _文(d.get("phone")).strip()
    住所 = _文(d.get("address")).strip()
    パス = _文(d.get("passcode") or d.get("password")).strip()

    if not 名:
        return {"status": "error", "message": "お名前を入力してください。"}
    if _漢字が要るか(名):
        return {"status": "error",
                "message": "お名前は漢字でご入力ください。"
                           "ひらがなでご登録いただくと、あとから同じ方が二重に登録されてしまうことがあります。"}
    if not 生年月日:
        return {"status": "error", "message": "生年月日を入力してください。"}
    if not _パスコードとして正しいか(パス):
        return {"status": "error", "message": "パスコードは数字4桁または6桁で入力してください。"}

    with transaction.atomic():
        全員 = list(Member.objects.select_for_update().exclude(deleted=True))

        # ① お名前＋生年月日で照らす
        同じ = [m for m in 全員
                if _名前で照らす(m.name) == 名
                and (m.birthday.isoformat() if m.birthday else "") == 生年月日]

        if len(同じ) > 1:
            return {"status": "error",
                    "message": "同じお名前と生年月日の登録が複数あります。受付にお申し出ください。"}

        if len(同じ) == 1:
            m = 同じ[0]
            if m.passcode_hash:
                return {"status": "error",
                        "message": "すでに登録済みです。パスコードでログインしてください。"}
            # **空いている項目だけ埋める。上書きしない。**
            if かな and not (m.kana or "").strip():
                m.kana = _保存するかな(かな)
            if 電話 and not (m.phone or "").strip():
                m.phone = _電話を整える(電話)
            if 住所 and not (m.address or "").strip():
                m.address = 住所
            m.passcode_hash = make_password(パス)
            m.last_online_at = timezone.now()
            m.save()
            return _入口の答え(m, claimed=True)

        # 同じお名前で生年月日が無い方がいれば、本人か確かめようがないので受付へ。
        if any(_名前で照らす(m.name) == 名 and not m.birthday for m in 全員):
            return {"status": "error",
                    "message": "同じお名前のご登録がありますが、生年月日が記録されていないため確認できません。"
                               "恐れ入りますが、受付にお申し出ください。"}

        # お名前の書き方だけが違う同じ方（フリガナ＋生年月日が同じ）を止める。
        if かな:
            かな照合 = _保存するかな(かな)
            if any(_保存するかな(m.kana) and _保存するかな(m.kana) == かな照合
                   and (m.birthday.isoformat() if m.birthday else "") == 生年月日 for m in 全員):
                return {"status": "error",
                        "message": "フリガナと生年月日が同じご登録がすでにあります。"
                                   "お名前の書き方だけが違う可能性があります。"
                                   "新しく登録するとこれまでの記録が分かれてしまうため、"
                                   "恐れ入りますが受付にお申し出ください。"}

        # ② 新しく作る。**ここから先は全部そろっていること**
        if not かな:
            return {"status": "error", "message": "フリガナを入力してください。"}
        if not 電話:
            return {"status": "error", "message": "電話番号を入力してください。"}
        if not 住所:
            return {"status": "error", "message": "ご住所を入力してください。"}

        使用中 = set(Member.objects.values_list("member_id", flat=True))
        会員ID = ""
        for _ in range(50):
            候補 = "MYM-" + str(random.randrange(1000, 10000))
            if 候補 not in 使用中:
                会員ID = 候補
                break
        if not 会員ID:
            # **黙って番号を作らない。**別の方の記録に重なるより、断るほうがよい。
            return {"status": "error",
                    "message": "会員番号を採番できませんでした。受付にお申し出ください。"}

        from django.utils.dateparse import parse_date

        今 = timezone.now()
        m = Member.objects.create(
            member_id=会員ID,
            created_at=今,
            name=_保存する名前(生の名),
            kana=_保存するかな(かな),
            phone=_電話を整える(電話),
            birthday=parse_date(生年月日),
            address=住所,
            passcode_hash=make_password(パス),
            device_sessions=[],
            # GAS 2487行 setUserRegistrationSource_(row, '新規登録', 'ランチャー')。分析の登録経路に出る。
            registration_source="新規登録",
            registration_source_detail="ランチャー",
            registration_source_updated_at=今,
            last_online_at=今,
        )
    return _入口の答え(m, claimed=False)


# ── ビジリスの利用登録 ───────────────────────────────

def ビジリス登録(d):
    """GAS の registerBijirisUse。ビジリスを使い始めた方に印を付ける。

    入口（start/index.html 925行）は **name / kana / passcode** を送ってくる。
    2026-09-15 まで memberId でしか探しておらず、切り替えると全員止まるところだった。
    memberId で来る古い呼び方も残す。
    """
    会員ID = _文(d.get("memberId")).strip()
    if 会員ID:
        with transaction.atomic():
            m = Member.objects.select_for_update().filter(pk=会員ID).first()
            if not m or m.deleted:
                return {"status": "error", "message": "会員情報が見つかりませんでした。"}
            m.bijiris_registered = True
            m.save(update_fields=["bijiris_registered", "changed_at"])
        return {"status": "ok", "memberId": 会員ID, "bijiris": True}

    名 = _名前で照らす(_文(d.get("name")))
    かな = _文(d.get("kana")).strip()
    パス = _文(d.get("passcode") or d.get("password")).strip()
    if not 名 or not かな or not パス:
        return {"status": "error", "message": "お名前・フリガナ・パスコードを入力してください。"}
    with transaction.atomic():
        当たり = [m for m in Member.objects.select_for_update().exclude(deleted=True)
                if _名前で照らす(m.name) == 名 and _パスコードが合う(m, パス)]
        if len(当たり) != 1:
            return {"status": "error", "message": "お名前またはパスコードが違います。"}
        m = 当たり[0]
        # **真偽値の列。**GAS はシートに「登録済み」と文字を入れているが、サーバーは BooleanField。
        m.bijiris_registered = True
        # フリガナが未登録の会員は、この機会に埋めておく（GAS 2619行）。
        if not (m.kana or "").strip():
            m.kana = _保存するかな(かな)
        m.last_online_at = timezone.now()
        m.save(update_fields=["bijiris_registered", "kana", "last_online_at", "changed_at"])
    return _入口の答え(m)
