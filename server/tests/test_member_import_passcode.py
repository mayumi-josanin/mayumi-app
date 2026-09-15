"""会員の取り込み: 表でパスコードの先頭の0を戻したあと、**他の項目が変わっていなくても**新しいハッシュが保存されること。

2026-09-15 の通し稽古で、22名のハッシュが古いまま残り「照合できない」になった。
"""

import json

import pytest
from django.contrib.auth.hashers import check_password, make_password
from django.core.management import call_command

from apps.members.models import Member

pytestmark = pytest.mark.django_db


def _書き出し(tmp_path, passcode):
    行 = {"memberId": "MYM-1001", "name": "佐藤花子", "kana": "さとうはなこ", "phone": "09012345678",
         "birthday": "1990-04-05", "address": "平塚市", "passcode": passcode, "createdAt": "2026-01-01 10:00",
         "stampCount": 0, "stampCardNumber": 1, "deviceSessions": "[]", "stampHistory": "[]", "rewardHistory": "[]"}
    p = tmp_path / "members.json"
    p.write_text(json.dumps({"書き出した日時": "2026-09-15", "件数": 1, "members": [行]}, ensure_ascii=False), encoding="utf-8")
    return str(p)


def test_パスコードだけ変わった会員も新しいハッシュで保存される(tmp_path):
    call_command("会員を取り込む", _書き出し(tmp_path, "123"))
    m = Member.objects.get(pk="MYM-1001")
    assert check_password("123", m.passcode_hash)
    # 表で 0 を戻した（他の項目は同じ）
    call_command("会員を取り込む", _書き出し(tmp_path, "0123"))
    m.refresh_from_db()
    assert check_password("0123", m.passcode_hash) and not check_password("123", m.passcode_hash)


def test_合っているハッシュは作り直さない(tmp_path):
    Member.objects.create(member_id="MYM-1001", name="佐藤花子", passcode_hash=make_password("0123"))
    前 = Member.objects.get(pk="MYM-1001").passcode_hash
    call_command("会員を取り込む", _書き出し(tmp_path, "0123"))
    assert Member.objects.get(pk="MYM-1001").passcode_hash == 前
