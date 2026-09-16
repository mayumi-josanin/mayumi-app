"""ビジリス管理「特典」。**中身は旧ビジリス管理アプリ（#page-gacha）と同じ**
（マイルストーン特典・月別ガチャ特典・キャンペーンスタンプ・受け取り状況）。"""

from datetime import datetime

import pytest
from django.utils import timezone

from apps.bijiris import gate, preferences
from apps.bijiris.models import CustomerProfile, Response
from apps.records.models import AppSetting

pytestmark = pytest.mark.django_db

URL = "/manage/bijiris/rewards/"


def 日本時間(*a):
    return timezone.make_aware(datetime(*a))


def _回答(name, at, response_id, plan, sheet, round_):
    return Response.objects.create(response_id=response_id, customer_name=name, submitted_at=at, survey_key="survey_bijiris_session",
                                   survey_title="施術後アンケート", answers=[
                                       {"questionId": "q_bijiris_session_ticket_plan", "value": plan},
                                       {"questionId": "q_bijiris_session_ticket_sheet", "value": sheet},
                                       {"questionId": "q_bijiris_session_ticket_round", "value": round_}])


def _設定():
    return AppSetting.objects.get(pk="bijiris_preferences").value


def test_設定が無いときの画面(as_owner):
    page = as_owner.get(URL).content.decode()
    for s in ["<h1>特典管理</h1>", "マイルストーン特典", "回数券を使い切った「完了枚数」がしきい値に達したお客様に渡す特典を設定できます。",
              "お客様アプリに表示する", "しきい値(枚)", "特典内容説明", "行を追加", "月別ガチャ特典",
              "行ごとに月、列ごとに A〜D賞の特典内容と出現確率を設定できます。", "月を追加", "A賞", "D賞", "出現確率 (%)",
              "キャンペーンスタンプ", "お客様アプリにキャンペーンスタンプを表示する", "受け取り状況", "特典の節目に届いたお客様はまだいません。"]:
        assert s in page, s
    # 既定: 表示オン・空の行1つ・今月のガチャ1行（25%ずつ）・キャンペーンオン
    assert 'id="ms-enabled" name="enabled" checked' in page and 'name="threshold_0"' in page and 'name="threshold_1"' not in page
    今月 = preferences.月の鍵(timezone.localdate().isoformat())
    assert f'name="month_0" class="form-control gacha-month" value="{今月}"' in page and 'value="25"' in page and "合計 100%" in page
    assert 'id="cp-enabled" name="enabled" checked' in page
    assert gate.断る文() in page and 'class="active">特典</a>' in page


def test_印が立つまでは保存できない(as_owner):
    for 経路, 入力 in (("milestone", {"row": ["0"], "threshold_0": "3", "reward_0": "お茶"}),
                   ("gacha", {"row": ["0"], "month_0": "2026-10"}), ("campaign", {})):
        r = as_owner.post(URL + 経路 + "/", 入力, follow=True)
        assert gate.断る文() in r.content.decode(), 経路
    assert not AppSetting.objects.filter(pk="bijiris_preferences").exists()


def test_マイルストーン特典の保存は正規化を通る(as_owner):
    gate.切り替える("server")
    r = as_owner.post(URL + "milestone/", {
        "row": ["0", "1", "2", "3"],
        "threshold_0": "10", "reward_0": "施術1回", "description_0": "10枚で",
        "threshold_1": "3", "reward_1": " お茶 ", "description_1": "",
        "threshold_2": "", "reward_2": "しきい値なし", "description_2": "",
        "threshold_3": "5", "reward_3": "", "description_3": "内容なし",
    }, follow=True)
    html = r.content.decode()
    assert "マイルストーン特典設定を保存しました。" in html
    assert _設定()["milestoneRewardConfig"] == {"enabled": False, "milestones": [
        {"threshold": 3, "reward": "お茶", "description": ""}, {"threshold": 10, "reward": "施術1回", "description": "10枚で"}]}
    # 他の設定は既定で埋まったまま（updatePreferences_ と同じ）
    assert _設定()["campaignStampEnabled"] is True and _設定()["notificationSubject"] == preferences.既定の通知件名
    assert 'value="3"' in html and 'value="お茶"' in html and 'id="ms-enabled" name="enabled">' in html
    # 表示オンで保存
    as_owner.post(URL + "milestone/", {"enabled": "on", "row": ["0"], "threshold_0": "3", "reward_0": "お茶"})
    assert _設定()["milestoneRewardConfig"]["enabled"] is True


def test_月別ガチャ特典の保存と言い分(as_owner):
    gate.切り替える("server")
    html = as_owner.post(URL + "gacha/", {"row": ["0"], "month_0": ""}, follow=True).content.decode()
    assert "月を選択してください。" in html
    html = as_owner.post(URL + "gacha/", {"row": ["0", "1"], "month_0": "2026-10", "month_1": "2026/10"}, follow=True).content.decode()
    assert "同じ月が重複しています。月ごとに1行だけ設定してください。" in html
    html = as_owner.post(URL + "gacha/", {}, follow=True).content.decode()
    assert "設定する月は1件以上必要です。" in html
    assert not AppSetting.objects.filter(pk="bijiris_preferences").exists()
    r = as_owner.post(URL + "gacha/", {
        "row": ["0", "1"],
        "month_1": "2026-09", "content_1_A": "A賞", "probability_1_A": "101", "content_1_B": "B賞", "probability_1_B": "abc",
        "month_0": "2026-11", "content_0_A": " 11月A ", "probability_0_A": "40.54", "probability_0_D": "60",
    }, follow=True)
    html = r.content.decode()
    assert "ガチャ特典設定を保存しました。" in html
    月々 = _設定()["gachaPrizeConfig"]["monthlyPrizes"]
    assert [m["month"] for m in 月々] == ["2026-09", "2026-11"]
    assert 月々[0]["prizes"]["A"] == {"content": "A賞", "probability": 100} and 月々[0]["prizes"]["B"]["probability"] == 0
    assert 月々[1]["prizes"]["A"] == {"content": "11月A", "probability": 40.5} and 月々[1]["prizes"]["D"]["probability"] == 60
    assert "2026年9月" in html and "2026年11月" in html and "合計 100%" in html and "合計 100.5%" in html and "badge-red" in html
    # 次に足す月はいちばん遅い月の翌月
    assert "var month = '2026-12'" in html


def test_キャンペーンスタンプの保存(as_owner):
    gate.切り替える("server")
    html = as_owner.post(URL + "campaign/", {}, follow=True).content.decode()
    assert "キャンペーンスタンプの設定を保存しました。" in html and _設定()["campaignStampEnabled"] is False
    assert 'id="cp-enabled" name="enabled">' in html
    as_owner.post(URL + "campaign/", {"enabled": "on"})
    assert _設定()["campaignStampEnabled"] is True


def test_受け取り状況は節目に届いた方だけ未受取が多い順(as_owner):
    AppSetting.objects.create(pk="bijiris_preferences", value={"milestoneRewardConfig": {"milestones": [
        {"threshold": 1, "reward": "お茶"}, {"threshold": 2, "reward": "施術1回"}]}})
    # 花子: 6回券を1枚使い切り（1個）＋手当て1 = 2個。1個は渡した
    CustomerProfile.objects.create(name="佐藤花子", member_number="MYM-0012", ticket_stamp_adjustment=1,
                                   reward_redemptions={"1": {"handed": True, "handedAt": "2026-08-20T00:00:00.000Z"}})
    for i in range(1, 7):
        _回答("佐藤花子", 日本時間(2026, 5, i, 10, 0), f"h{i}", "6回券", "1枚目", f"{i}回目")
    # 太郎: プロフィール無し。10回券を使い切って1個。未受取1
    for i in range(1, 11):
        _回答("山田太郎", 日本時間(2026, 6, i, 10, 0), f"t{i}", "10回券", "1枚目", f"{i}回目")
    # 次郎: 3回目まで。届いていない
    _回答("鈴木次郎", 日本時間(2026, 6, 1, 10, 0), "j1", "10回券", "1枚目", "3回目")
    page = as_owner.get(URL).content.decode()
    assert "鈴木次郎" not in page
    assert page.find(">佐藤花子</a>") < page.find(">山田太郎</a>")
    行 = page[page.find(">佐藤花子</a>"):page.find(">山田太郎</a>")]
    assert "MYM-0012" in 行 and "2個</td>" in 行 and "未受取 1件" in 行 and "1件</td>" in 行
    assert "1個 お茶" in 行 and "受取済み" in 行 and "2026/08/20" in 行 and "2個 施術1回" in 行
    行 = page[page.find(">山田太郎</a>"):]
    assert "1個</td>" in 行 and "未受取 1件" in 行 and "0件</td>" in 行
    # 「渡した」の付け外しは顧客管理で
    assert "/manage/bijiris/customers/%E4%BD%90%E8%97%A4%E8%8A%B1%E5%AD%90/" in page
