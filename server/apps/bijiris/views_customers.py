"""ビジリス管理「🧾 顧客管理」。**中身は旧ビジリス管理アプリ（#page-customers）と同じ。**

    顧客一覧      検索（会員番号・氏名・フリガナ）・回数券・回答状態の絞り込み。新しい回答がある方から上に
    顧客情報      お名前と会員番号・通知状態・回答数・回数券の枚数・キャンペーン・最新回答と最新測定・回数券スタンプ
    顧客情報を編集 パスコードの再設定を許可・お名前／会員番号／回数券の種類・何枚目・何回目・顧客を削除
    顧客メモ      記録日つきのメモ（新しい順）
    マイルストーン特典 スタンプ（施術後アンケートから＋手当て）・−1／＋1・道のり・特典の受け取り（渡した）
    測定履歴      最新測定日・WHR・左右差・部位ごとの初回比、測定ごとの前回比と初回比（見るだけ）
    アンケート回答履歴 アンケートごとの回答数と最新、その中の回答の一覧

数え方は旧アプリの app.js（renderCustomerManagement 4385行〜）と Code.gs を写している。
**スタンプの数は2通りある**（旧アプリもそう）:
  - 「回数券 合計 N枚」（顧客情報のカード）は app.js の getCompletedTicketCardBreakdown。
    計測時アンケートの「回数券終了時」でもそのカードを満了とみなす
  - 「マイルストーン特典」のスタンプは Code.gs の getCompletedTicketCardCountsByCustomer_。
    施術後アンケートで最終回まで出したカードだけ数え、それに手当て（ticketStampAdjustment）を足す。
    **お客様のアプリに出るのはこちら。**渡す判断はこの数で行う

## お名前で会員を探さない（CLAUDE.md）

顧客の鍵はお名前（CustomerProfile.name）。会員番号（member_number）はビジリス GAS が写した記録で、
`member_id` は**受付が画面で「まゆみの会員と結ぶ」を選んだときだけ**入る。お名前が同じでも自動では結ばない。

## 書き込みは切り替え（9/24 以降）まで断る

gate.ビジリスはサーバーが正() が立つまで、保存・削除・メモ・スタンプ・受け取り・パスコード許可は
すべて gate.断る文() を出して一覧へ戻す（読むのはよい）。
旧アプリで削除すると回答・写真・測定・メモ・プロフィールが全部消える（deleteCustomerProfile_）。
ここでも同じ範囲を消す。写真のファイル（media/bijiris/）は回答の記録と一緒に消える。

## 外したもの

- 測定データの追加・編集・削除・推移グラフ・CSV／PDF出力（計測は一覧だけ。編集は別の画面で）
- 時系列カルテ・PDF出力・回答の印刷・写真保存・回答の削除（回答管理の画面で）
- 回答の詳細（回答管理へ）
"""

import re
from datetime import datetime, timedelta

from django.contrib import messages
from django.db import transaction
from django.http import Http404
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required
from apps.measurements.models import Measurement
from apps.members.models import Member

from . import gate, preferences
from .models import CustomerProfile, Response

# 旧アプリの TICKET_INFO_QUESTION_IDS（Code.gs の CUSTOMER_TICKET_INFO_QUESTION_IDS と同じ）
回数券の設問 = {
    "plan": ["q_bijiris_session_ticket_plan", "q_ticket_end_ticket_size"],
    "sheet": ["q_bijiris_session_ticket_sheet", "q_ticket_end_ticket_sheet"],
    "round": ["q_bijiris_session_ticket_round", "q_ticket_end_ticket_round"],
}
回数券の種類 = ["6回券", "10回券"]
何枚目の選択肢 = [f"{i}枚目" for i in range(1, 21)]
何回目の選択肢 = [f"{i}回目" for i in range(1, 11)]
計測の部位 = [("waist", "ウエスト"), ("hip", "ヒップ"), ("thigh_right", "太もも右"), ("thigh_left", "太もも左")]
パスコード再設定の持ち時間 = timedelta(minutes=30)


# ---------------------------------------------------------------------------
# 文字・日付
# ---------------------------------------------------------------------------

def _文字(値) -> str:
    return "" if 値 is None else str(値).strip()


def _数字を拾う(値) -> int:
    """parseTicketLabelNumber: 「3枚目」→ 3。無ければ 0。"""
    m = re.search(r"\d+", _文字(値))
    return int(m.group(0)) if m else 0


def _回数券の回数(plan) -> int:
    """parseTicketCount: 10 を含めば 10、6 を含めば 6、それ以外は最初の数字。"""
    s = _文字(plan)
    if "10" in s:
        return 10
    if "6" in s:
        return 6
    return _数字を拾う(s)


def 会員番号を整える(値) -> str:
    """normalizeMemberNumber_: 数字を拾って MYM-0000 の形に。数字が無ければ空。"""
    n = _数字を拾う(_文字(値).upper())
    return f"MYM-{n:04d}" if n > 0 else ""


def _日時(値):
    """ISO 文字列や datetime を日本時間に。読めなければ None。"""
    if not 値:
        return None
    if hasattr(値, "tzinfo"):
        t = 値
    else:
        from django.utils.dateparse import parse_datetime

        t = parse_datetime(_文字(値).replace("Z", "+00:00"))
        if t is None:
            return None
    if timezone.is_naive(t):
        t = timezone.make_aware(t)
    return timezone.localtime(t)


def 日時の表示(値) -> str:
    """旧アプリの formatDate（2026/08/10 10:05）。"""
    t = _日時(値)
    return t.strftime("%Y/%m/%d %H:%M") if t else "-"


def 日付の表示(値) -> str:
    """旧アプリの formatDateOnly（2026/08/10）。date でもよい。"""
    if hasattr(値, "strftime") and not hasattr(値, "tzinfo"):
        return 値.strftime("%Y/%m/%d")
    t = _日時(値)
    return t.strftime("%Y/%m/%d") if t else "-"


def _スタンプの日(値) -> str:
    """旧アプリの formatStampDate（26/08/10）。"""
    t = _日時(値)
    return t.strftime("%y/%m/%d") if t else ""


def _今日() -> str:
    return timezone.localdate().strftime("%Y-%m-%d")


def _メモの日(値) -> str:
    """normalizeMemoDate_: YYYY-MM-DD だけ残す。読めなければ空。"""
    s = _文字(値)
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    t = _日時(s)
    return t.strftime("%Y-%m-%d") if t else ""


# ---------------------------------------------------------------------------
# 回答から回数券の情報を読む（app.js getResponseTicketInfo など）
# ---------------------------------------------------------------------------

def _回答の値(answers, 設問の候補) -> str:
    """getAnswerValueByQuestionIds_: 候補の設問のうち最初に値があるもの。"""
    for a in answers or []:
        if isinstance(a, dict) and a.get("questionId") in 設問の候補:
            v = a.get("value")
            v = "".join(_文字(x) for x in v) if isinstance(v, list) else _文字(v)
            if v:
                return v
    return ""


def _回答の文(answers, question_id) -> str:
    for a in answers or []:
        if isinstance(a, dict) and a.get("questionId") == question_id:
            v = a.get("value")
            return "".join(_文字(x) for x in v) if isinstance(v, list) else _文字(v)
    return ""


def 回答の回数券(r) -> dict:
    """{plan, sheet, round}（文字のまま。無ければ空）。"""
    return {k: _回答の値(r.answers, 回数券の設問[k]) for k in ("plan", "sheet", "round")}


def _回数券がある(r) -> bool:
    return any(回答の回数券(r).values())


def _回数券終了の回答か(r) -> bool:
    return "回数券終了" in _回答の文(r.answers, "q_measure_timing")


def _生きている回答(name):
    """ゴミ箱以外。新しい順（getActiveCustomerResponses）。"""
    return list(Response.objects.filter(customer_name=name).exclude(status="trash").order_by("-submitted_at", "-sheet_row"))


def 現在の回数券(profile, 回答一覧) -> dict | None:
    """app.js getCurrentTicketInfoForCustomer。
    プロフィールの回数券カードが正。無ければ回答から。計測時「回数券終了時」がその後にあれば満了に。
    返す形: {plan, sheet_number, round, total, sheet_label, round_label}"""
    card = profile.active_ticket_card if profile and isinstance(profile.active_ticket_card, dict) else None
    最新の回数券回答 = next((r for r in 回答一覧 if _回数券がある(r)), None)
    if card and _文字(card.get("plan")):
        plan = _文字(card.get("plan"))
        sheet_number = max(1, int(card.get("sheetNumber") or 0))
        round_ = max(0, int(card.get("round") or 0))
    elif 最新の回数券回答:
        info = 回答の回数券(最新の回数券回答)
        plan = info["plan"]
        if not plan:
            return None
        sheet_number = _数字を拾う(info["sheet"]) or 1
        round_ = _数字を拾う(info["round"])
    else:
        return None
    total = _回数券の回数(plan)
    回答のカード = _数字を拾う(回答の回数券(最新の回数券回答)["sheet"]) if 最新の回数券回答 else 0
    if total and sheet_number == 回答のカード and 最新の回数券回答:
        終了 = next((r for r in 回答一覧 if _回数券終了の回答か(r)), None)
        if 終了 and 終了.submitted_at and 最新の回数券回答.submitted_at and 終了.submitted_at >= 最新の回数券回答.submitted_at:
            round_ = total
    return {"plan": plan, "sheet_number": sheet_number, "round": round_, "total": total,
            "sheet_label": f"{sheet_number}枚目", "round_label": f"{round_}回目"}


def _スタンプの日付(回数券, 回答一覧) -> dict:
    """getCustomerTicketStampDates: 何回目 → その回答の送信日時。同じ種類・同じ枚目の回答だけ。"""
    出 = {}
    if not 回数券 or not 回数券["total"]:
        return 出
    for r in 回答一覧:
        info = 回答の回数券(r)
        round_ = _数字を拾う(info["round"])
        if not round_ or round_ > 回数券["total"] or round_ in 出:
            continue
        if 回数券["plan"] and info["plan"] != 回数券["plan"]:
            continue
        if 回数券["sheet_label"] and info["sheet"] != 回数券["sheet_label"]:
            continue
        出[round_] = r.submitted_at
    return 出


def 回数券のスタンプ(回数券, 回答一覧) -> list:
    """回数券スタンプの丸（1回目〜）。押した回には日付を添える。"""
    if not 回数券 or not 回数券["total"]:
        return []
    日付 = _スタンプの日付(回数券, 回答一覧)
    round_ = max(0, min(回数券["total"], 回数券["round"]))
    return [{"step": i, "active": i <= round_, "date": _スタンプの日(日付.get(i)) if i <= round_ else ""}
            for i in range(1, 回数券["total"] + 1)]


def 満了したカードの内訳(回答一覧) -> dict:
    """app.js getCompletedTicketCardBreakdown（顧客情報の「回数券 合計 N枚」）。
    計測時アンケートの「回数券終了時」でも、その直前の回数券のカードを満了とみなす。"""
    最終回 = {}
    for r in 回答一覧:
        info = 回答の回数券(r)
        plan, sheet, round_ = info["plan"], _数字を拾う(info["sheet"]), _数字を拾う(info["round"])
        if not plan or sheet <= 0 or round_ <= 0:
            continue
        鍵 = f"{plan}|{sheet}"
        if 鍵 not in 最終回 or round_ > 最終回[鍵]["round"]:
            最終回[鍵] = {"plan": plan, "round": round_}
    古い順 = sorted((r for r in 回答一覧 if 回答の回数券(r)["plan"] and r.submitted_at), key=lambda r: r.submitted_at)
    for 終了 in (r for r in 回答一覧 if _回数券終了の回答か(r) and r.submitted_at):
        直前 = [r for r in 古い順 if r.submitted_at <= 終了.submitted_at]
        if not 直前:
            continue
        info = 回答の回数券(直前[-1])
        plan, sheet = info["plan"], _数字を拾う(info["sheet"])
        if not plan or sheet <= 0:
            continue
        上限 = _回数券の回数(plan)
        if 上限 > 0:
            最終回[f"{plan}|{sheet}"] = {"plan": plan, "round": 上限}
    合計, 種類別 = 0, {}
    for card in 最終回.values():
        上限 = _回数券の回数(card["plan"])
        if 上限 > 0 and card["round"] >= 上限:
            合計 += 1
            種類別[card["plan"]] = 種類別.get(card["plan"], 0) + 1
    return {"total": 合計, "by_plan": 種類別}


def 満了したカードの数(回答一覧) -> int:
    """Code.gs getCompletedTicketCardCountsByCustomer_（マイルストーン特典のスタンプ）。
    施術後アンケートで最終回まで出したカードだけ。計測時の「回数券終了時」は数えない。"""
    最終回 = {}
    for r in 回答一覧:
        info = 回答の回数券(r)
        plan, sheet, round_ = info["plan"], _数字を拾う(info["sheet"]), _数字を拾う(info["round"])
        if not plan or sheet <= 0 or round_ <= 0:
            continue
        鍵 = f"{plan}|{sheet}"
        if 鍵 not in 最終回 or round_ > 最終回[鍵]["round"]:
            最終回[鍵] = {"plan": plan, "round": round_}
    return sum(1 for c in 最終回.values() if _回数券の回数(c["plan"]) > 0 and c["round"] >= _回数券の回数(c["plan"]))


def キャンペーン(回答一覧) -> dict:
    """app.js computeCampaignCards: 施術内容「キャンペーン」で1回、計測時「キャンペーン終了」で1枚完了。"""
    出来事 = []
    for r in 回答一覧:
        if not r.submitted_at:
            continue
        if _回答の文(r.answers, "q_bijiris_session_type") == "キャンペーン":
            出来事.append((r.submitted_at, "visit"))
        elif "キャンペーン終了" in _回答の文(r.answers, "q_measure_timing"):
            出来事.append((r.submitted_at, "close"))
    完了, 進行中 = 0, 0
    for _, kind in sorted(出来事, key=lambda x: x[0]):
        if kind == "visit":
            進行中 += 1
        else:
            完了 += 1
            進行中 = 0
    return {"completed": 完了, "current": 進行中}


def 通知の状態(push_status) -> dict:
    """app.js describePushStatus。"""
    s = push_status if isinstance(push_status, dict) else None
    if not s or not any(k in s for k in ("enabled", "supported", "permission")):
        return {"label": "未設定", "detail": "顧客側でまだ通知設定が同期されていません。", "badge": "badge-gray", "updated": ""}
    更新 = 日時の表示(s.get("updatedAt")) if s.get("updatedAt") else ""
    if s.get("supported") is not True:
        return {"label": "未対応", "detail": "この端末または表示方法では通知を利用できません。", "badge": "badge-gray", "updated": 更新}
    if s.get("enabled") is True:
        return {"label": "オン", "detail": "顧客アプリで通知受信が有効です。", "badge": "badge-green", "updated": 更新}
    if _文字(s.get("permission")).lower() == "denied":
        return {"label": "拒否", "detail": "端末側で通知が拒否されています。", "badge": "badge-red", "updated": 更新}
    return {"label": "オフ", "detail": "顧客アプリで通知はオフです。", "badge": "badge-gray", "updated": 更新}


# ---------------------------------------------------------------------------
# マイルストーン特典（app.js renderCustomerMilestoneSection）
# ---------------------------------------------------------------------------

def マイルストーンの状況(profile, 回答一覧, 設定=None) -> dict:
    設定 = 設定 or preferences.読む()["milestoneRewardConfig"]
    アンケートから = 満了したカードの数(回答一覧)
    手当て = int(profile.ticket_stamp_adjustment or 0) if profile else 0
    合計 = max(0, アンケートから + 手当て)
    受け取り = (profile.reward_redemptions if profile and isinstance(profile.reward_redemptions, dict) else {}) or {}
    行 = []
    未受取 = 0
    for m in 設定["milestones"]:
        達成 = 合計 >= m["threshold"]
        r = 受け取り.get(str(m["threshold"])) or {}
        渡した = r.get("handed") is True
        if 達成 and not 渡した:
            未受取 += 1
        行.append({**m, "achieved": 達成, "handed": 渡した,
                   "handed_at": 日時の表示(r.get("handedAt")) if 渡した and r.get("handedAt") else "",
                   "status": ("未達成", "badge-gray") if not 達成 else (("受取済み", "badge-blue") if 渡した else ("未受取", "badge-green"))})
    到達点 = 設定["milestones"][-1]["threshold"] if 設定["milestones"] else 10
    特典の節目 = {m["threshold"] for m in 設定["milestones"]}
    道のり = [{"step": i, "reached": 合計 >= i, "reward": i in 特典の節目} for i in range(1, 到達点 + 1)]
    return {"enabled": 設定["enabled"], "total": 合計, "from_surveys": アンケートから, "adjustment": 手当て,
            "adjustment_text": f"{'+' if 手当て >= 0 else ''}{手当て}", "rows": 行, "pending": 未受取,
            "goal": 到達点, "road": 道のり, "has_milestones": bool(設定["milestones"])}


# ---------------------------------------------------------------------------
# 測定履歴（app.js buildMeasurementHistoryRows / renderMeasurementSummaryCards）
# ---------------------------------------------------------------------------

def _差(いま, 前):
    if いま is None or 前 is None:
        return None
    return round(float(いま) - float(前), 1)


def _cm(値) -> str:
    return "-" if 値 is None else f"{float(値):.1f}cm"


def _差の表示(値) -> dict:
    """{text, cls}。プラスは increase、マイナスは decrease。"""
    if 値 is None:
        return {"text": "-", "cls": "neutral"}
    if 値 == 0:
        return {"text": "0.0cm", "cls": "neutral"}
    return {"text": f"{'+' if 値 > 0 else ''}{値:.1f}cm", "cls": "increase" if 値 > 0 else "decrease"}


def _whr(値) -> str:
    return "-" if 値 is None else f"{float(値):.3f}"


def 測定の一覧(name) -> dict:
    """まとめ（最新測定日・最新 WHR・左右差・部位ごとの初回比）と履歴（新しい順。前回比・初回比）。"""
    古い順 = list(Measurement.objects.filter(customer_name=name).order_by("measured_on", "created_at"))
    最初 = 古い順[0] if 古い順 else None
    履歴 = []
    for i, m in enumerate(古い順):
        前 = 古い順[i - 1] if i > 0 else None
        部位 = []
        for key, label in 計測の部位:
            v = getattr(m, key)
            部位.append({"label": label, "value": _cm(v),
                       "prev": _差の表示(_差(v, getattr(前, key)) if 前 else None),
                       "first": _差の表示(_差(v, getattr(最初, key)) if 最初 else None)})
        左右差 = _差(m.thigh_right, m.thigh_left)
        履歴.append({"date": 日付の表示(m.measured_on), "metrics": 部位, "whr": _whr(m.whr),
                   "gap": "-" if 左右差 is None else f"{abs(左右差):.1f}cm", "memo": m.staff_memo or "-"})
    履歴.reverse()
    最新 = 古い順[-1] if 古い順 else None
    左右差 = _差(最新.thigh_right, 最新.thigh_left) if 最新 else None
    まとめ = [
        {"label": "最新測定日", "value": 日付の表示(最新.measured_on) if 最新 else "-", "meta": f"履歴 {len(古い順)}件"},
        {"label": "最新 WHR", "value": _whr(最新.whr) if 最新 else "-", "meta": f"初回 {_whr(最初.whr) if 最初 else '-'}"},
        {"label": "太もも左右差", "value": "-" if 左右差 is None else f"{abs(左右差):.1f}cm", "meta": "左右のバランス確認用"},
    ]
    for key, label in 計測の部位:
        v = getattr(最新, key) if 最新 else None
        差 = _差の表示(_差(v, getattr(最初, key)) if 最初 else None)
        まとめ.append({"label": label, "value": _cm(v), "meta": f"初回比 {差['text']}", "delta_cls": 差["cls"]})
    return {"summary": まとめ, "rows": 履歴, "latest": 最新}


# ---------------------------------------------------------------------------
# 顧客一覧（app.js getCustomerDirectoryEntries / getFilteredCustomerDirectoryEntries）
# ---------------------------------------------------------------------------

def _顧客名簿() -> list:
    """顧客プロフィール ∪ ゴミ箱以外の回答のお名前。
    並びは旧アプリと同じ latest_at（プロフィールの更新日時と回答の送信日時の新しい方）。回答が無い方は会員番号・お名前順。
    表に出す「最新回答」は回答の送信日時だけ（latest_response_at）。"""
    名簿 = {}
    for p in CustomerProfile.objects.all():
        名簿[p.name] = {"name": p.name, "profile": p, "member_number": p.member_number, "latest_at": p.updated_at,
                      "latest_response_at": None, "count": 0, "has_responses": False}
    for r in Response.objects.exclude(status="trash").order_by("-submitted_at"):
        name = _文字(r.customer_name)
        if not name:
            continue
        e = 名簿.setdefault(name, {"name": name, "profile": None, "member_number": "", "latest_at": r.submitted_at,
                               "latest_response_at": None, "count": 0, "has_responses": False})
        e["count"] += 1
        e["has_responses"] = True
        if r.submitted_at and (not e["latest_at"] or r.submitted_at > e["latest_at"]):
            e["latest_at"] = r.submitted_at
        if r.submitted_at and (not e["latest_response_at"] or r.submitted_at > e["latest_response_at"]):
            e["latest_response_at"] = r.submitted_at
    最古 = timezone.make_aware(datetime(1970, 1, 1))
    return sorted(名簿.values(), key=lambda e: ((0, -(e["latest_at"] or 最古).timestamp()) if e["latest_at"] else (1, 0),
                                              e["member_number"], e["name"]))


def _整理(値) -> str:
    return _文字(値).lower()


@owner_required
def customer_list(request):
    q = _整理(request.GET.get("q"))
    ticket = _文字(request.GET.get("ticket"))
    state = _文字(request.GET.get("state"))
    全員 = _顧客名簿()
    shown = []
    for e in 全員:
        回答一覧 = _生きている回答(e["name"]) if (q or ticket) else []
        回数券 = 現在の回数券(e["profile"], 回答一覧) if (q or ticket) else None
        if q:
            材料 = " ".join(x for x in [e["name"], e["member_number"], e["profile"].name_kana if e["profile"] else "",
                                    " ".join([回数券["plan"], 回数券["sheet_label"], 回数券["round_label"]]) if 回数券 else ""] if x)
            if q not in 材料.lower():
                continue
        if ticket and not (回数券 and 回数券["plan"] == ticket):
            continue
        if state == "with" and not e["has_responses"]:
            continue
        if state == "without" and e["has_responses"]:
            continue
        shown.append({**e, "latest": 日時の表示(e["latest_response_at"]) if e["latest_response_at"] else "-",
                      "kana": e["profile"].name_kana if e["profile"] else ""})
    会員 = list(Member.objects.filter(deleted=False).only("member_id", "name").order_by("member_id"))
    return render(request, "bijiris/customer_list.html", {
        "items": shown, "q": request.GET.get("q", ""), "ticket": ticket, "state": state, "ticket_plans": 回数券の種類,
        "total": len(全員), "members": 会員,
        "server_is_source": gate.ビジリスはサーバーが正(), "refuse_text": gate.断る文(),
    })


# ---------------------------------------------------------------------------
# 顧客情報（詳細）
# ---------------------------------------------------------------------------

def _顧客を引く(name):
    """プロフィールが無くても回答があれば顧客。どちらも無ければ 404。"""
    name = _文字(name)
    profile = CustomerProfile.objects.filter(name=name).first()
    回答一覧 = _生きている回答(name)
    if not profile and not Response.objects.filter(customer_name=name).exists():
        raise Http404("顧客が見つかりません。")
    return name, profile, 回答一覧


def _回答の見出し(r) -> str:
    """formatResponseEntryTitle: 「回数券 10回券 / 2枚目 / 3回目」、無ければアンケート名。"""
    info = 回答の回数券(r)
    ある = [("回数券", info["plan"]), ("何枚目", info["sheet"]), ("何回目", info["round"])]
    ある = [(l, v) for l, v in ある if v]
    if ある:
        return " / ".join(f"{l} {v}" if i == 0 else v for i, (l, v) in enumerate(ある))
    return r.survey_title or "回答詳細"


def _アンケート別(回答一覧) -> list:
    """groupResponsesBySurvey: アンケートごとに回答数・最新・回答の一覧（新しい順）。最新が新しい順。"""
    群 = {}
    for r in 回答一覧:
        鍵 = r.survey_key or r.survey_title or r.response_id
        g = 群.setdefault(鍵, {"key": 鍵, "title": r.survey_title or "アンケート", "latest_at": r.submitted_at, "count": 0, "responses": []})
        g["count"] += 1
        if r.submitted_at and (not g["latest_at"] or r.submitted_at > g["latest_at"]):
            g["latest_at"] = r.submitted_at
        info = 回答の回数券(r)
        g["responses"].append({"id": r.pk, "title": _回答の見出し(r), "date": 日時の表示(r.submitted_at),
                              "status": r.get_status_display(), "status_cls": {"new": "badge-orange", "checked": "badge-blue", "done": "badge-green"}.get(r.status, "badge-gray"),
                              "ticket": [(l, v) for l, v in (("回数券", info["plan"]), ("何枚目", info["sheet"]), ("何回目", info["round"])) if v],
                              "photos": r.photos.count()})
    最古 = timezone.make_aware(datetime(1970, 1, 1))
    出 = sorted(群.values(), key=lambda g: g["latest_at"] or 最古, reverse=True)
    for g in 出:
        g["latest"] = 日時の表示(g["latest_at"])
    return 出


@owner_required
def customer_detail(request, name):
    name, profile, 回答一覧 = _顧客を引く(name)
    回数券 = 現在の回数券(profile, 回答一覧)
    測定 = 測定の一覧(name)
    内訳 = 満了したカードの内訳(回答一覧)
    最新回答 = 回答一覧[0] if 回答一覧 else None
    通知 = 通知の状態(profile.push_status if profile else None)
    会員 = Member.objects.filter(pk=profile.member_id).first() if profile and profile.member_id else None
    メモ = [{"at": 日付の表示(e.get("at")) if e.get("at") else "-", "memo": e.get("memo", "")}
          for e in (profile.memo_entries if profile else []) if isinstance(e, dict) and e.get("memo")]
    return render(request, "bijiris/customer_detail.html", {
        "name": name, "profile": profile,
        "title_name": f"{profile.member_number} / {name}" if profile and profile.member_number else name,
        "kana": (profile.name_kana if profile else "") or "-",
        "push": 通知,
        "response_count": len(回答一覧), "survey_count": len({r.survey_key or r.survey_title or r.response_id for r in 回答一覧}),
        "cards": 内訳, "cards_6": 内訳["by_plan"].get("6回券", 0), "cards_10": 内訳["by_plan"].get("10回券", 0),
        "campaign": キャンペーン(回答一覧),
        "latest_response": f"{最新回答.survey_title} / {日時の表示(最新回答.submitted_at)}" if 最新回答 else "-",
        "latest_measurement": f"{日付の表示(測定['latest'].measured_on)} / WHR {_whr(測定['latest'].whr)}" if 測定["latest"] else "-",
        "ticket": 回数券, "stamps": 回数券のスタンプ(回数券, 回答一覧),
        "ticket_plans": 回数券の種類, "sheet_options": 何枚目の選択肢, "round_options": 何回目の選択肢,
        "draft": {"member_number": profile.member_number if profile else "",
                  "plan": 回数券["plan"] if 回数券 else "", "sheet": 回数券["sheet_label"] if 回数券 else "",
                  "round": 回数券["round_label"] if 回数券 else ""},
        "linked_member": 会員,
        "members": Member.objects.filter(deleted=False).only("member_id", "name").order_by("member_id"),
        "today": _今日(), "memos": メモ,
        "milestone": マイルストーンの状況(profile, 回答一覧),
        "measurement": 測定,
        "survey_groups": _アンケート別(回答一覧),
        "server_is_source": gate.ビジリスはサーバーが正(), "refuse_text": gate.断る文(),
    })


# ---------------------------------------------------------------------------
# 書き込み（印が立つまで断る）
# ---------------------------------------------------------------------------

def _断る(request, name=None):
    messages.error(request, gate.断る文())
    return redirect("manage:bijiris:customer_detail", name=name) if name else redirect("manage:bijiris:customer_list")


def _プロフィールを用意する(name):
    """無ければ作る（旧アプリの updateAdminCustomerProfileRecord_ は回答しか無い方にも作る）。"""
    p, _ = CustomerProfile.objects.get_or_create(name=name)
    return p


@owner_required
@require_POST
def customer_add(request):
    """顧客を1名足す（旧アプリには無い。回答が無いお客様の回数券カードやメモを先に持てるように）。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request)
    name = _文字(request.POST.get("name"))
    if not name:
        messages.error(request, "お名前を入力してください。")
        return redirect("manage:bijiris:customer_list")
    if CustomerProfile.objects.filter(name=name).exists():
        messages.error(request, f"{name} 様はすでに登録されています。")
        return redirect("manage:bijiris:customer_detail", name=name)
    CustomerProfile.objects.create(
        name=name, name_kana=re.sub(r"\s+", "", _文字(request.POST.get("name_kana"))),
        member_number=会員番号を整える(request.POST.get("member_number")), admin_managed=True, updated_at=timezone.now())
    messages.success(request, "顧客を追加しました。")
    return redirect("manage:bijiris:customer_detail", name=name)


@owner_required
@require_POST
def customer_save(request, name):
    """旧アプリの saveCustomerProfile → Code.gs updateCustomerProfile_。
    お名前を変えると回答・測定・メモも付いていく。回数券は最新の回数券回答の答えも書き換える。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, 回答一覧 = _顧客を引く(name)
    次の名 = _文字(request.POST.get("name"))
    会員番号 = 会員番号を整える(request.POST.get("member_number"))
    plan, sheet, round_ = (_文字(request.POST.get(k)) for k in ("ticket_plan", "ticket_sheet", "ticket_round"))
    if not 次の名:
        messages.error(request, "お名前を入力してください。")
        return redirect("manage:bijiris:customer_detail", name=name)
    回数券を直す = bool(plan or sheet or round_)
    if 回数券を直す and not (plan and sheet and round_):
        messages.error(request, "回数券情報は種類・何枚目・何回目をすべて選択してください。")
        return redirect("manage:bijiris:customer_detail", name=name)
    if 次の名 != name and CustomerProfile.objects.filter(name=次の名).exists():
        messages.error(request, f"{次の名} 様はすでに別の顧客として登録されています。")
        return redirect("manage:bijiris:customer_detail", name=name)
    with transaction.atomic():
        今 = timezone.now()
        if 回数券を直す:
            最新 = next((r for r in 回答一覧 if _回数券がある(r)), None)
            if 最新:
                answers = list(最新.answers or [])
                for 候補, 値 in ((回数券の設問["plan"], plan), (回数券の設問["sheet"], sheet), (回数券の設問["round"], round_)):
                    for a in answers:
                        if isinstance(a, dict) and a.get("questionId") in 候補:
                            a["value"] = 値
                最新.answers = answers
                最新.managed_at = 今
                最新.save(update_fields=["answers", "managed_at", "changed_at"])
        if 次の名 != name:
            Response.objects.filter(customer_name=name).update(customer_name=次の名, managed_at=今)
            Measurement.objects.filter(customer_name=name).update(customer_name=次の名)
        p = profile or _プロフィールを用意する(name)
        if 次の名 != name:
            別名 = [a for a in (p.aliases or []) if a != 次の名]
            if name not in 別名:
                別名.append(name)
            p.aliases = 別名
            p.name = 次の名
        if 会員番号:
            p.member_number = 会員番号
        if 回数券を直す:
            p.active_ticket_card = {"plan": plan, "sheetNumber": _数字を拾う(sheet), "round": _数字を拾う(round_)}
            p.active_ticket_card_source = "admin"
        p.admin_managed = True
        p.updated_at = 今
        p.save()
        if p.member_number:
            Measurement.objects.filter(customer_name=p.name).update(member_number=p.member_number)
    messages.success(request, "顧客情報を保存しました。")
    return redirect("manage:bijiris:customer_detail", name=次の名)


@owner_required
@require_POST
def customer_link_member(request, name):
    """まゆみの会員と結ぶ。**受付が会員番号を選んだときだけ** member_id を入れる（お名前では当てない）。空にすると外す。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, _ = _顧客を引く(name)
    member_id = _文字(request.POST.get("member_id"))
    if member_id and not Member.objects.filter(pk=member_id, deleted=False).exists():
        messages.error(request, "その会員番号の会員が見つかりません。")
        return redirect("manage:bijiris:customer_detail", name=name)
    p = profile or _プロフィールを用意する(name)
    p.member_id = member_id
    p.updated_at = timezone.now()
    p.save(update_fields=["member_id", "updated_at", "changed_at"])
    messages.success(request, f"まゆみの会員 {member_id} と結びました。" if member_id else "まゆみの会員との結びつきを外しました。")
    return redirect("manage:bijiris:customer_detail", name=name)


@owner_required
@require_POST
def customer_delete(request, name):
    """旧アプリの deleteCustomerProfile → Code.gs deleteCustomerProfile_。回答（写真も）・測定・メモ・プロフィールを全部消す。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, _ = _顧客を引く(name)
    with transaction.atomic():
        回答数 = Response.objects.filter(customer_name=name).count()
        Response.objects.filter(customer_name=name).delete()
        Measurement.objects.filter(customer_name=name).delete()
        if profile:
            profile.delete()
        # 別名にこのお名前を持つ方からも外す（deleteCustomerProfileRecord_ と同じ）
        for p in CustomerProfile.objects.exclude(aliases=[]):
            if name in (p.aliases or []):
                p.aliases = [a for a in p.aliases if a != name]
                p.save(update_fields=["aliases", "changed_at"])
    messages.success(request, f"顧客情報を削除しました。（回答 {回答数}件も削除）")
    return redirect("manage:bijiris:customer_list")


@owner_required
@require_POST
def customer_passcode_setup(request, name):
    """旧アプリの allowPasscodeSetup → Code.gs allowPasscodeSetup_。今のパスコードを消し、30分だけ再設定できるようにする。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, _ = _顧客を引く(name)
    if not profile:
        messages.error(request, "お客様が見つかりませんでした。")
        return redirect("manage:bijiris:customer_detail", name=name)
    今 = timezone.now()
    profile.passcode_hash = ""
    profile.passcode_salt = ""
    profile.passcode_setup_until = (今 + パスコード再設定の持ち時間).isoformat().replace("+00:00", "Z")
    profile.updated_at = 今
    profile.save(update_fields=["passcode_hash", "passcode_salt", "passcode_setup_until", "updated_at", "changed_at"])
    messages.success(request, "30分間、お客様がパスコードを設定できます。アプリの「パスコードを忘れた方」からお進みください。")
    return redirect("manage:bijiris:customer_detail", name=name)


@owner_required
@require_POST
def customer_memo(request, name):
    """旧アプリの saveCustomerMemo → Code.gs updateCustomerMemo_。同じ日の同じメモは重ねない。新しい順に100件まで。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, _ = _顧客を引く(name)
    at = _メモの日(request.POST.get("at"))
    memo = _文字(request.POST.get("memo"))
    if not at:
        messages.error(request, "記録日を入力してください。")
        return redirect("manage:bijiris:customer_detail", name=name)
    if not memo:
        messages.error(request, "メモを入力してください。")
        return redirect("manage:bijiris:customer_detail", name=name)
    p = profile or _プロフィールを用意する(name)
    entries = [e for e in (p.memo_entries or []) if isinstance(e, dict) and _文字(e.get("memo"))]
    if not any(_文字(e.get("memo")) == memo and _メモの日(e.get("at")) == at for e in entries):
        entries.insert(0, {"at": at, "memo": memo})
    entries.sort(key=lambda e: _メモの日(e.get("at")), reverse=True)
    p.memo_entries = entries[:100]
    p.latest_memo = p.memo_entries[0]["memo"] if p.memo_entries else ""
    p.updated_at = timezone.now()
    p.save(update_fields=["memo_entries", "latest_memo", "updated_at", "changed_at"])
    messages.success(request, "顧客メモを保存しました。")
    return redirect("manage:bijiris:customer_detail", name=name)


@owner_required
@require_POST
def customer_stamp_adjust(request, name):
    """旧アプリの adjustTicketStamp: 手当てを ±1。合計がマイナスになる引き算はしない。幅は -50〜50（Code.gs）。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, 回答一覧 = _顧客を引く(name)
    try:
        delta = int(request.POST.get("delta") or 0)
    except ValueError:
        delta = 0
    if delta not in (1, -1):
        messages.error(request, "スタンプは1個ずつ足したり戻したりします。")
        return redirect("manage:bijiris:customer_detail", name=name)
    p = profile or _プロフィールを用意する(name)
    次 = int(p.ticket_stamp_adjustment or 0) + delta
    if 満了したカードの数(回答一覧) + 次 < 0:
        messages.error(request, "これ以上は減らせません。")
        return redirect("manage:bijiris:customer_detail", name=name)
    p.ticket_stamp_adjustment = max(-50, min(50, 次))
    p.updated_at = timezone.now()
    p.save(update_fields=["ticket_stamp_adjustment", "updated_at", "changed_at"])
    messages.success(request, "スタンプを1個足しました。" if delta > 0 else "スタンプを1個戻しました。")
    return redirect("manage:bijiris:customer_detail", name=name)


@owner_required
@require_POST
def customer_redemption(request, name):
    """旧アプリの updateRewardRedemption → Code.gs updateCustomerRewardRedemption_。「渡した」の付け外し。"""
    if not gate.ビジリスはサーバーが正():
        return _断る(request, name)
    name, profile, _ = _顧客を引く(name)
    try:
        threshold = int(request.POST.get("threshold") or 0)
    except ValueError:
        threshold = 0
    if threshold <= 0:
        messages.error(request, "しきい値が正しくありません。")
        return redirect("manage:bijiris:customer_detail", name=name)
    if not profile:
        messages.error(request, "お客様が見つかりません。")
        return redirect("manage:bijiris:customer_detail", name=name)
    handed = request.POST.get("handed") == "1"
    受け取り = dict(profile.reward_redemptions or {})
    if handed:
        受け取り[str(threshold)] = {"handed": True, "handedAt": timezone.now().isoformat().replace("+00:00", "Z")}
    else:
        受け取り.pop(str(threshold), None)
    profile.reward_redemptions = 受け取り or None
    profile.updated_at = timezone.now()
    profile.save(update_fields=["reward_redemptions", "updated_at", "changed_at"])
    messages.success(request, "渡し済みにしました。" if handed else "未渡しに戻しました。")
    return redirect("manage:bijiris:customer_detail", name=name)
