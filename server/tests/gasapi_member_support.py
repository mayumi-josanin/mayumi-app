"""会員の窓口（/api）を叩く試験の、共通の道具。

test_gasapi_member*.py から使う。**ここに試験は書かない**（pytest が拾わないよう
test_ で始めていない）。

## なぜ Django の test client で「窓口」を叩くのか

会員の24窓口は、GAS が `SERVER_TABLES` に `member` を入れた瞬間に
`サーバーへ渡す.js` 経由で **そのまま** ここへ届く。関数を直接呼ぶ試験だと、
合鍵の要否・POST の形・JSON の往復（真偽値・null・文字列の "false"）が
試験から抜ける。**GAS が送ってくる形で叩く。**
"""

import json

from django.contrib.auth.hashers import make_password

from apps.members.models import Member

# 管理者向けの窓口に付ける合鍵。GAS は SERVER_API_KEY を X-Api-Key で送ってくる。
合鍵 = "api-key-for-member-tests"


def 読む(client, action, 引数=None, 合鍵あり=False):
    """GAS の サーバーから読む_ と同じ形。`?action=...&data=<JSON>`。"""
    q = {"action": action}
    if 引数 is not None:
        q["data"] = json.dumps(引数, ensure_ascii=False)
    extra = {"HTTP_X_API_KEY": 合鍵} if 合鍵あり else {}
    return client.get("/api", q, **extra)


def 書く(client, 中身, 合鍵あり=False):
    """GAS の サーバーへ書く_ と同じ形。POST の JSON に `type` が入る。"""
    extra = {"HTTP_X_API_KEY": 合鍵} if 合鍵あり else {}
    return client.post("/api", data=json.dumps(中身, ensure_ascii=False),
                       content_type="application/json", **extra)


def 会員を作る(member_id="MYM-1001", name="佐藤花子", passcode="1234", **項目):
    """会員を1人作る。**パスコードはサーバーの作法どおりハッシュにして入れる。**

    entrance.py / writes.py は `check_password` で照合するので、平文を入れると
    誰もログインできない試験になる。
    """
    if passcode:
        項目["passcode_hash"] = make_password(passcode)
    項目.setdefault("device_sessions", [])
    return Member.objects.create(member_id=member_id, name=name, **項目)


def 読み直す(m):
    return Member.objects.get(pk=m.pk)
