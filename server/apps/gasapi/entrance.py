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

import hashlib
import random
import re
import unicodedata

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

    当たり = []
    for m in Member.objects.exclude(deleted=True):
        if _名前で照らす(m.name) != 名:
            continue
        if m.passcode_hash and check_password(パス, m.passcode_hash):
            当たり.append(m)
        elif m.password_hash and m.password_salt:
            # 旧いパスワードの形（ハッシュ＋ソルト）。GAS と同じ作り方でそろえる。
            種 = (パス + m.password_salt).encode("utf-8")
            if hashlib.sha256(種).hexdigest() == m.password_hash:
                当たり.append(m)

    # **ちょうど1人でなければ通さない。**
    # 0人は「違う」、2人以上は同姓同名で見分けがつかない。
    if len(当たり) != 1:
        _はずれを数える(名)
        return {"status": "error", "message": "お名前またはパスコードが違います。"}

    m = 当たり[0]
    _数えるのをやめる(名)
    m.last_online_at = timezone.now()
    m.save(update_fields=["last_online_at", "changed_at"])
    return {"status": "ok", "memberId": m.member_id, "name": m.name or "", "kana": m.kana or ""}


# ── 新規登録 ─────────────────────────────────────────

def 新規登録(d):
    """GAS の registerAccount。

    **すでにある記録を引き受ける場合がある。**受付が先に作った行に、
    お客様があとからパスコードを付ける形。お名前と生年月日で照らす。
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
    if not 生年月日:
        return {"status": "error", "message": "生年月日を入力してください。"}
    if not _パスコードとして正しいか(パス):
        return {"status": "error", "message": "パスコードは数字4桁または6桁で入力してください。"}

    with transaction.atomic():
        # ① お名前＋生年月日で照らす
        同じ = [m for m in Member.objects.select_for_update().exclude(deleted=True)
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
                m.kana = かな
            if 電話 and not (m.phone or "").strip():
                m.phone = _電話を整える(電話)
            if 住所 and not (m.address or "").strip():
                m.address = 住所
            m.passcode_hash = make_password(パス)
            m.save()
            return {"status": "ok", "memberId": m.member_id, "name": m.name or "",
                    "kana": m.kana or "", "claimed": True}

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

        m = Member.objects.create(
            member_id=会員ID,
            created_at=timezone.now(),
            name=生の名,
            kana=かな,
            phone=_電話を整える(電話),
            birthday=parse_date(生年月日),
            address=住所,
            passcode_hash=make_password(パス),
            device_sessions=[],
        )
    return {"status": "ok", "memberId": m.member_id, "name": m.name or "",
            "kana": m.kana or "", "claimed": False}


# ── ビジリスの利用登録 ───────────────────────────────

def ビジリス登録(d):
    """GAS の registerBijirisUse。ビジリスを使い始めた方に印を付ける。"""
    会員ID = _文(d.get("memberId")).strip()
    if not 会員ID:
        return {"status": "error", "message": "会員IDが指定されていません。"}
    with transaction.atomic():
        m = Member.objects.select_for_update().filter(pk=会員ID).first()
        if not m or m.deleted:
            return {"status": "error", "message": "会員情報が見つかりませんでした。"}
        # **真偽値の列。**GAS はシートに「登録済み」と文字を入れているが、
        # サーバー側は BooleanField（2026-09-05 に試験で判明）。
        # 文字を入れようとして落ちた。**列の型を見てから書く。**
        m.bijiris_registered = True
        m.save(update_fields=["bijiris_registered", "changed_at"])
    return {"status": "ok", "memberId": 会員ID, "bijiris": True}
