"""会員の24窓口が、サーバーで **動くこと** と、合鍵の要否。

GAS（gas/管理者・お客様.js）は `SERVER_TABLES` に `member` が入ると、
`サーバーへ渡すか_('member')` を見ている24か所から、ここへ転送してくる。
**1つでも notImplemented を返すと、GAS はそれを「サーバーに無い」と扱い、
書き込みは「サーバーに届きませんでした」でお客様に返る。**

個々の窓口の中身（返す項目・文言）は

    test_gasapi_member_entrance.py   入口（loginAccount / registerAccount / registerBijirisUse）
    test_gasapi_member_customer.py   お客様側（updateUser / 端末 / 特典 / 復元 …）
    test_gasapi_member_admin.py      管理側（getAdminUsers / updateAdminUser / mergeUsers …）

ここでは「全部ある」「合鍵の要否が GAS と同じ」だけを固定する。
"""

import re
from pathlib import Path

import pytest

from apps.gasapi import views
from tests.gasapi_member_support import 会員を作る, 書く, 読む

pytestmark = [pytest.mark.django_db, pytest.mark.usefixtures("会員試験の下ごしらえ")]


# GAS の24か所（2026-09-15 時点）。読む窓口は `?action=`、書く窓口は POST の `type`。
#
#   checkMemberToken は札の検証を GAS に残し、会員を見に行く部分だけ
#   `checkMemberOnServer` として来る（管理者・お客様.js 2269行）。
#   getAnalyticsData（6668行）は分析の窓口。会員から作る2項目を
#   サーバーが埋める前提で、GAS が差し替えをやめる。
読む窓口 = [
    "getUserDevices", "getUserRewardStatus", "getRecoveryCandidates",
    "getAdminUsers", "getPushUsers", "checkMemberOnServer", "getAnalytics",
]
書く窓口 = [
    "updateUser", "recoverAccount", "resetForgottenPasscode",
    "syncUserRewardStatus", "drawRewardGacha", "unsubscribePush",
    "syncUserDeviceSession", "removeUserDeviceSession",
    "loginAccount", "registerAccount", "registerBijirisUse",
    "updateAdminUser", "deleteUser", "updateAdminRewardStatus",
    "grantSurveyStamp", "issueTransferCode", "mergeUsers",
]
# GAS には無いが、移行後にシートを直せない代わりに管理アプリが直接呼ぶ2つ
# （admin/index.html 7449行・7482行）。
サーバーだけの窓口 = ["setUserPasscode", "stopUserPush"]


def test_24か所ぶんの窓口がそろっている():
    assert len(読む窓口) + len(書く窓口) == 24


@pytest.mark.parametrize("action", 読む窓口)
def test_読む窓口がnotImplementedを返さない(client, action):
    会員を作る()
    r = 読む(client, action, {"memberId": "MYM-1001"}, 合鍵あり=True)
    assert r.status_code == 200, r.content
    assert not r.json().get("notImplemented"), r.json()


@pytest.mark.parametrize("type_", 書く窓口 + サーバーだけの窓口)
def test_書く窓口がnotImplementedを返さない(client, type_):
    # 中身が足りなくてもよい。「まだ無い」と答えないことだけを見る。
    r = 書く(client, {"type": type_}, 合鍵あり=True)
    assert r.status_code == 200, r.content
    assert not r.json().get("notImplemented"), r.json()
    # GAS はエラーでも 200 で status を返す。形がそろっていること。
    assert r.json().get("status") in ("ok", "error")


def test_知らない窓口はnotImplementedと答える(client):
    """黙って成功にしない。GAS 側が「サーバーに無い」と気づけるように。"""
    r = 書く(client, {"type": "somethingNew"}, 合鍵あり=True)
    assert r.json()["notImplemented"] is True
    r = 読む(client, "getSomethingNew", 合鍵あり=True)
    assert r.status_code == 501 and r.json()["notImplemented"] is True


# ---------------------------------------------------------------------------
# 合鍵の要否
#
# お客様アプリ・入口は合鍵を持っていない（GitHub Pages で公開されているため）。
# GAS の PUBLIC_ACTIONS（1946行）と同じものは合鍵なしで通り、
# それ以外（管理アプリ向け）は合鍵が要る。
# ---------------------------------------------------------------------------

お客様の読む窓口 = ["getUserDevices", "getUserRewardStatus", "getRecoveryCandidates"]
お客様の書く窓口 = [
    "updateUser", "recoverAccount", "issueTransferCode", "resetForgottenPasscode",
    "syncUserDeviceSession", "removeUserDeviceSession", "unsubscribePush",
    "syncUserRewardStatus", "drawRewardGacha",
    "loginAccount", "registerAccount", "registerBijirisUse",
]
管理の読む窓口 = ["getAdminUsers", "getPushUsers", "checkMemberOnServer", "getAnalytics"]
管理の書く窓口 = ["updateAdminUser", "deleteUser", "updateAdminRewardStatus",
                  "grantSurveyStamp", "mergeUsers", "setUserPasscode", "stopUserPush"]


@pytest.mark.parametrize("action", お客様の読む窓口)
def test_お客様の読む窓口は合鍵なしで通る(client, action):
    r = 読む(client, action, {"memberId": "MYM-1001"})
    assert r.status_code == 200


@pytest.mark.parametrize("type_", お客様の書く窓口)
def test_お客様の書く窓口は合鍵なしで通る(client, type_):
    r = 書く(client, {"type": type_})
    assert r.status_code == 200


@pytest.mark.parametrize("action", 管理の読む窓口)
def test_管理の読む窓口は合鍵が要る(client, action):
    assert 読む(client, action, {"memberId": "MYM-1001"}).status_code == 403
    assert 読む(client, action, {"memberId": "MYM-1001"}, 合鍵あり=True).status_code == 200


@pytest.mark.parametrize("type_", 管理の書く窓口)
def test_管理の書く窓口は合鍵が要る(client, type_):
    assert 書く(client, {"type": type_}).status_code == 403
    assert 書く(client, {"type": type_}, 合鍵あり=True).status_code == 200


def test_合鍵が未設定なら管理の窓口は全部止まる(client, settings):
    """設定し忘れて素通しになる事故を防ぐ。お客様の窓口は影響を受けない。"""
    settings.API_KEY = ""
    assert 読む(client, "getAdminUsers").status_code == 503
    assert 読む(client, "getUserDevices", {"memberId": "x"}).status_code == 200


def test_公開アクションはGASのPUBLIC_ACTIONSと同じ():
    """ずれると、切り替えた瞬間にお客様の操作が 403 になる（CLAUDE.md）。"""
    gas = Path(__file__).resolve().parents[2] / "gas" / "管理者・お客様.js"
    if not gas.exists():
        pytest.skip("GAS のファイルが無い環境")
    文 = gas.read_text(encoding="utf-8")
    先頭 = 文.index("const PUBLIC_ACTIONS = {")
    末尾 = 文.index("};", 先頭)
    GAS側 = set(re.findall(r"(\w+):\s*true", 文[先頭:末尾]))
    # 入口の3つは GAS の PUBLIC_ACTIONS には無いが、doPost（2816〜2825行）が
    # requireAdminAccess_ より前に通している。**サーバーでも合鍵なしで通す。**
    入口 = {"loginAccount", "registerAccount", "registerBijirisUse"}
    assert views.公開アクション == GAS側 | 入口


# ---------------------------------------------------------------------------
# 24か所目: getAnalyticsData（6668行）
#
# GAS は order を渡しているとき、サーバーの分析を使うが、会員から作る
# `registrationRoutes` と `categoryUsage` だけは **member を渡すまで** GAS で
# 作って差し替える。member を渡した瞬間に差し替えをやめるので、
# **サーバーが埋めていないと分析画面の2項目が空になる。**
# ---------------------------------------------------------------------------

def test_分析の登録経路は会員の表から作る(client):
    from django.utils import timezone

    会員を作る("MYM-1001", registration_source="新規登録", created_at=timezone.now())
    会員を作る("MYM-1002", registration_source="復元", created_at=timezone.now(),
            registration_source_updated_at=timezone.now())
    中 = 読む(client, "getAnalytics", 合鍵あり=True).json()
    経路 = 中["registrationRoutes"]
    assert 経路["routeLabels"] == ["新規登録", "復元", "引き継ぎコード利用", "重複候補からの復旧", "自動復旧"]
    assert 経路["totals"]["新規登録"] == 1
    assert 経路["totals"]["復元"] == 1
    assert len(経路["months"]) == 1
