"""会員の取り込み: **サーバーにたまった LINEID を、古い表で流し直しても消さない。**

9/19 の切り替えより後は、シートではなくサーバーが正になる。
LINEID はそのあと、お客様アプリや予約システムから**サーバー側にたまっていく**。
そこで「念のため」もう一度取り込みを流すと、シートに無いぶんが空に上書きされ、
その方はLINEとつながらなくなる。気づくのは、お客様が「つながりません」とおっしゃったとき。
"""

import json

import pytest
from django.core.management import call_command

from apps.members.models import Member

pytestmark = pytest.mark.django_db


def _書き出し(tmp_path, line_id=None):
    行 = {"memberId": "MYM-1001", "name": "佐藤花子", "kana": "さとうはなこ", "phone": "09012345678",
          "birthday": "1990-04-05", "address": "平塚市", "createdAt": "2026-01-01 10:00",
          "stampCount": 0, "stampCardNumber": 1, "deviceSessions": "[]", "stampHistory": "[]",
          "rewardHistory": "[]"}
    if line_id is not None:
        行["lineUserId"] = line_id
    p = tmp_path / "members.json"
    p.write_text(json.dumps({"書き出した日時": "2026-09-19", "件数": 1, "members": [行]}, ensure_ascii=False),
                 encoding="utf-8")
    return str(p)


def test_表にLINEIDが無くてもサーバーのものは消さない(tmp_path):
    call_command("会員を取り込む", _書き出し(tmp_path))
    # 切り替えのあと、お客様アプリや予約システムからサーバーに入った
    Member.objects.filter(pk="MYM-1001").update(line_user_id="U1234567890abcdef")

    call_command("会員を取り込む", _書き出し(tmp_path))  # 古い表で流し直す

    assert Member.objects.get(pk="MYM-1001").line_user_id == "U1234567890abcdef"


def test_表にLINEIDがあれば今までどおり上書きする(tmp_path):
    call_command("会員を取り込む", _書き出し(tmp_path, "Uあたらしい"))
    assert Member.objects.get(pk="MYM-1001").line_user_id == "Uあたらしい"

    call_command("会員を取り込む", _書き出し(tmp_path, "Uなおした"))
    assert Member.objects.get(pk="MYM-1001").line_user_id == "Uなおした"


def test_どちらにも無ければ空のまま(tmp_path):
    call_command("会員を取り込む", _書き出し(tmp_path))
    assert Member.objects.get(pk="MYM-1001").line_user_id is None


def test_消さなかった会員は変わったことにしない(tmp_path, capsys):
    """LINEID を残しただけで「中身が変わる」に数えない（当日の突き合わせが読みにくくなる）。"""
    call_command("会員を取り込む", _書き出し(tmp_path))
    Member.objects.filter(pk="MYM-1001").update(line_user_id="U1234567890abcdef")
    capsys.readouterr()

    call_command("会員を取り込む", _書き出し(tmp_path))

    出た = capsys.readouterr().out
    assert "変わらない:   1名" in 出た
