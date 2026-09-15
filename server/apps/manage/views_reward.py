"""スタンプ・特典管理。**中身は旧管理アプリ（#page-stamp-rewards）と同じ。**

    🎁 月別ガチャ特典設定       月を足す・保存（GAS の getRewardGachaConfig / saveRewardGachaConfig）
    🎟 会員別スタンプ・特典状況  検索・絞り込み・表・詳細・編集（GAS の getAdminUsers / updateAdminRewardStatus）

会員の状況は gasapi/admin_member.一覧()（GAS の getAdminUsers と同じ形）から作り、
編集の保存は admin_member.特典を書き換える()（GAS の handleUpdateAdminRewardStatus）へ渡す。
**新しい書き込み経路は作らない。**

## 「特典データから会員を再作成」ボタンは作らない

旧管理アプリにある `restoreUsersFromRewardData` は、GAS の doPost にもう
受け口が無く（`未定義のPOSTアクション` が返る）、押しても失敗する。
サーバーでは会員の表そのものが正なので、復元する元も無い。**外した。**

## 会員の表は 9/19 までスプレッドシートが正

会員ごとの編集（書き込み）は、`member_gate.会員はサーバーが正()` が立つまで断る。
月別ガチャ特典設定は会員の表ではない（AppSetting）ので、断らない。
"""

import re
from datetime import datetime, timedelta

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_datetime
from django.views.decorators.http import require_POST

from apps.gasapi import admin_member, writes
from apps.gasapi.views import _ガチャ設定
from apps.members.models import Member

from . import member_gate
from .permissions import owner_required

スタンプの上限 = 10
特典の既定名 = "スタンプ達成特典"
賞 = writes._ガチャの賞


# ---------------------------------------------------------------------------
# 日付の読み書き
#
# 特典の一覧（rewards）はシートから移した文字列（'2026-04-01T10:00:00+09:00' や
# '2026/04/01 10:00:00'）が混ざっている。**読めない値は空扱い**にして落とさない。
# ---------------------------------------------------------------------------

_日付だけ = re.compile(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$")


def _時刻に(値):
    """いろいろな書き方の日時を、日本時間の aware datetime に。読めなければ None。"""
    if isinstance(値, datetime):
        t = 値
    else:
        文 = str(値 or "").strip()
        if not 文:
            return None
        t = parse_datetime(文.replace(" ", "T", 1).replace("/", "-"))
        if t is None:
            m = _日付だけ.match(文.split("T")[0])
            if not m:
                return None
            try:
                t = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    if timezone.is_naive(t):
        t = timezone.make_aware(t)
    return timezone.localtime(t)


def _時刻の字(t):
    """GAS の formatDateTime_ と同じ形（'2026-04-01T10:00:00+09:00'）。"""
    return timezone.localtime(t).strftime("%Y-%m-%dT%H:%M:%S%z").replace("+0900", "+09:00") if t else ""


def _表示日(値):
    """旧アプリの formatDisplayDate と同じ（'2026/04/01'。空は 'ー'）。"""
    t = _時刻に(値)
    return t.strftime("%Y/%m/%d") if t else "ー"


def _入力用の日(値):
    t = _時刻に(値)
    return t.strftime("%Y-%m-%d") if t else ""


def _入力用の日時(値):
    t = _時刻に(値)
    return t.strftime("%Y-%m-%dT%H:%M") if t else ""


def _その日の始まり(文):
    日 = parse_date((文 or "").strip()[:10]) if 文 else None
    return _時刻の字(timezone.make_aware(datetime.combine(日, datetime.min.time()))) if 日 else ""


def _その日の終わり(文):
    日 = parse_date((文 or "").strip()[:10]) if 文 else None
    return _時刻の字(timezone.make_aware(datetime.combine(日, datetime.max.time().replace(microsecond=0)))) if 日 else ""


def _数(値, 既定=0):
    try:
        return int(float(値))
    except (TypeError, ValueError):
        return 既定


# ---------------------------------------------------------------------------
# 会員ごとの見方（旧アプリの getUserRewards / getRewardProgressInfo などと同じ）
# ---------------------------------------------------------------------------

def _特典一覧(user):
    """旧アプリの normalizeRewardList。"""
    出 = []
    for i, r in enumerate(user.get("rewards") or []):
        if not isinstance(r, dict):
            r = {}
        説明 = r.get("rewardNote") if r.get("rewardNote") is not None else r.get("note")
        出.append({
            "id": str(r.get("id") or f"reward-{i + 1}"),
            "rewardName": str(r.get("rewardName") or 特典の既定名),
            "rewardNote": str(説明 or ""),
            "cardNum": max(1, _数(r.get("cardNum"), 1) or 1),
            "earnedDate": r.get("earnedDate") or "",
            "expiryDate": r.get("expiryDate") or "",
            "used": r.get("used") is True,
            "usedAt": r.get("usedAt") or "",
        })
    return 出


def _期限切れ(r):
    t = _時刻に(r.get("expiryDate"))
    return bool(t) and t < timezone.localtime(timezone.now())


def _受け取り状況(r):
    if r["used"]:
        return {"label": "受取済み", "badge": "badge-green"}
    if _期限切れ(r):
        return {"label": "期限切れ", "badge": "badge-red"}
    return {"label": "未受取", "badge": "badge-blue"}


def _特典の数(rewards):
    return {
        "total": len(rewards),
        "unused": sum(1 for r in rewards if not r["used"] and not _期限切れ(r)),
        "used": sum(1 for r in rewards if r["used"]),
        "expired": sum(1 for r in rewards if not r["used"] and _期限切れ(r)),
    }


def _最新特典日(rewards):
    ある = [(t, r["earnedDate"]) for r in rewards if r["earnedDate"] and (t := _時刻に(r["earnedDate"]))]
    return max(ある)[1] if ある else ""


def _近い有効期限(rewards):
    """旧アプリの getClosestRewardExpiry: 未受取で期限内のいちばん近いもの → 期限切れの新しいもの → 受取済みの新しいもの。"""
    def 期限(r):
        return _時刻に(r["expiryDate"])

    生き = sorted((r for r in rewards if not r["used"] and 期限(r) and not _期限切れ(r)), key=期限)
    if 生き:
        return 生き[0]["expiryDate"]
    切れ = sorted((r for r in rewards if not r["used"] and 期限(r)), key=期限, reverse=True)
    if 切れ:
        return 切れ[0]["expiryDate"]
    済 = sorted((r for r in rewards if 期限(r)), key=期限, reverse=True)
    return 済[0]["expiryDate"] if 済 else ""


def _進み具合(stamp_count, unused):
    if stamp_count >= スタンプの上限:
        return {"label": "達成済み", "badge": "badge-orange"}
    if unused > 0:
        return {"label": "未受取特典あり", "badge": "badge-green"}
    if stamp_count > 0:
        return {"label": "収集中", "badge": "badge-blue"}
    return {"label": "未取得", "badge": "badge-gray"}


def _最新の動き(user):
    """旧アプリの getLatestStampActivityTimestamp。並び替えに使う。"""
    履歴 = [t for e in (user.get("stampHistory") or []) if isinstance(e, dict) and (t := _時刻に(e.get("acquiredDate")))]
    候補 = [user.get("latestStampActivityAt"), user.get("lastStampAt"), max(履歴) if 履歴 else None,
          user.get("lastStampDate"), user.get("stampAchievedDate"), _最新特典日(_特典一覧(user)), user.get("timestamp")]
    for v in 候補:
        t = _時刻に(v)
        if t:
            return t
    return None


def _重複の理由():
    """旧アプリの buildDuplicateUsersFromAdminUsers と同じ判定（電話番号・氏名＋生年月日・氏名）。"""
    出 = {}
    電話, 氏名生年, 氏名 = {}, {}, {}
    for m in Member.objects.exclude(deleted=True).only("member_id", "name", "phone", "birthday"):
        p = re.sub(r"\D+", "", m.phone or "")
        n = re.sub(r"[\s　]+", "", m.name or "")
        b = m.birthday.strftime("%Y/%m/%d") if m.birthday else ""
        if p:
            電話.setdefault(p, []).append(m.member_id)
        if n and b:
            氏名生年.setdefault(f"{n}|{b}", []).append(m.member_id)
        if n:
            氏名.setdefault(n, []).append(m.member_id)
    for 表, 理由 in ((電話, "電話番号が一致"), (氏名生年, "氏名と生年月日が一致"), (氏名, "氏名が一致")):
        for ids in 表.values():
            if len(ids) < 2:
                continue
            for i in ids:
                if 理由 not in 出.setdefault(i, []):
                    出[i].append(理由)
    return 出


def _一件(user, 重複):
    rewards = _特典一覧(user)
    数 = _特典の数(rewards)
    stamp = max(0, _数(user.get("stampCount")))
    card = max(1, _数(user.get("stampCardNum"), 1) or 1)
    return {
        "user": user, "member_id": user.get("memberId") or "", "name": user.get("name") or "未設定",
        "phone": user.get("phone") or "電話番号未設定", "card": card, "stamp": stamp,
        "last_stamp": _表示日(user.get("lastStampDate")), "counts": 数,
        "expiry": _表示日(_近い有効期限(rewards)), "latest_reward": _表示日(_最新特典日(rewards)),
        "progress": _進み具合(stamp, 数["unused"]), "duplicate": 重複.get(user.get("memberId") or ""),
        "rewards": rewards, "activity": _最新の動き(user),
    }


def _絞り文字(値):
    """旧アプリの normalizeFilterText（小文字・空白なし・長音や全角ハイフンは '-'）。"""
    return re.sub(r"[‐－―ー]", "-", re.sub(r"\s+", "", str(値 or "").lower()))


# ---------------------------------------------------------------------------
# 月別ガチャ特典設定
# ---------------------------------------------------------------------------

def _今月():
    return timezone.localdate().strftime("%Y-%m")


def _月の見出し(月):
    return f"{月[:4]}年{int(月[5:7])}月" if 月 else "未設定"


def _その月の設定(設定, 月):
    """旧アプリの getActiveRewardGachaConfigEntry: その月 → その月より前でいちばん新しいもの → 最後。"""
    月ごと = 設定["monthlyPrizes"]
    合致 = 直前 = None
    for e in 月ごと:
        if e["month"] == 月:
            合致 = e
        elif e["month"] <= 月:
            直前 = e
    return 合致 or 直前 or 月ごと[-1]


def _ガチャの画面(設定):
    行 = []
    for i, e in enumerate(設定["monthlyPrizes"]):
        合計 = round(sum(float(e["prizes"][k]["probability"]) for k in 賞), 1)
        行.append({"index": i, "month": e["month"], "label": _月の見出し(e["month"]),
                   "prizes": [{"key": k, **e["prizes"][k]} for k in 賞],
                   "total": int(合計) if 合計 == int(合計) else 合計, "total_ok": abs(合計 - 100) < 0.001})
    今月 = _今月()
    return {
        "gacha_rows": 行,
        "gacha_meta": f"現在月は {_月の見出し(今月)} です。抽選では {_月の見出し(_その月の設定(設定, 今月)['month'])} の設定を参照し、未設定月は直近の月を使用します。",
        "gacha_next_month": _次の月(設定["monthlyPrizes"]),
        "gacha_defaults": writes._ガチャの既定確率,
    }


def _次の月(月ごと):
    最後 = max((e["month"] for e in 月ごと), default=_今月())
    年, 月 = int(最後[:4]), int(最後[5:7])
    return f"{年 + (月 // 12):04d}-{(月 % 12) + 1:02d}"


def _ガチャの入力(request):
    """表の入力を GAS へ送る形（{monthlyPrizes: [{month, prizes: {A: {content, note, probability}}}]}）に。"""
    出 = []
    for i in request.POST.getlist("row"):
        出.append({
            "month": request.POST.get(f"month_{i}", ""),
            "prizes": {k: {
                "content": request.POST.get(f"content_{i}_{k}", ""),
                "note": request.POST.get(f"note_{i}_{k}", ""),
                "probability": request.POST.get(f"probability_{i}_{k}", ""),
            } for k in 賞},
        })
    return {"monthlyPrizes": 出}


def _ガチャを確かめる(設定):
    """旧アプリの validateRewardGachaConfig と同じ言い分。問題が無ければ ''。"""
    月ごと = 設定.get("monthlyPrizes") or []
    if not 月ごと:
        return "設定する月がありません。"
    見た = set()
    for e in 月ごと:
        月 = writes._ガチャの月キー(e.get("month"))
        if not 月:
            return "対象月を選択してください。"
        if 月 in 見た:
            return "同じ月が重複しています。月ごとに1行だけ設定してください。"
        見た.add(月)
        合計 = sum(writes._ガチャの確率((e.get("prizes") or {}).get(k, {}).get("probability"), 0) for k in 賞)
        if abs(合計 - 100) > 0.001:
            return f"{_月の見出し(月)} の出現確率合計を 100% にしてください。"
    return ""


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------

@owner_required
def reward_list(request, gacha_input=None):
    q = (request.GET.get("q") or "").strip()
    stamp = request.GET.get("stamp", "all")
    ticket = request.GET.get("ticket", "all")

    重複 = _重複の理由()
    全員 = [_一件(u, 重複) for u in admin_member.一覧()["users"]]
    絞り = _絞り文字(q)
    shown = []
    for it in 全員:
        c, s = it["counts"], it["stamp"]
        if stamp == "completed" and s < スタンプの上限:
            continue
        if stamp == "collecting" and (s <= 0 or s >= スタンプの上限):
            continue
        if stamp == "empty" and s != 0:
            continue
        if ticket == "unused" and c["unused"] == 0:
            continue
        if ticket == "earned" and c["total"] == 0:
            continue
        if ticket == "none" and c["total"] > 0:
            continue
        u = it["user"]
        if 絞り and 絞り not in _絞り文字(" ".join(str(x or "") for x in [
                u.get("memberId"), u.get("name"), u.get("kana"), u.get("phone"), u.get("address"), u.get("memo"),
                u.get("lastStampDate"), _最新特典日(it["rewards"])])):
            continue
        shown.append(it)
    # 最終スタンプ取得日時が新しい会員から上に（同じなら登録日時の新しい順、会員ID順）
    最古 = timezone.make_aware(datetime(1970, 1, 1))
    shown.sort(key=lambda it: (-(it["activity"] or 最古).timestamp(), -(_時刻に(it["user"].get("timestamp")) or 最古).timestamp(), it["member_id"]))

    summary = [
        ("会員総数", len(全員)),
        ("達成済みカード", sum(1 for it in 全員 if it["stamp"] >= スタンプの上限)),
        ("未受取特典数", sum(it["counts"]["unused"] for it in 全員)),
        ("現在表示中", len(shown)),
    ]
    # 保存で弾かれたときは、入力した中身をそのまま見せる（打ち直させない）
    設定 = writes.ガチャ設定を整える(gacha_input, 確率を並べ替える=False) if gacha_input else \
        writes.ガチャ設定を整える(_ガチャ設定()["config"], 確率を並べ替える=False)
    return render(request, "manage/reward_list.html", {
        "items": shown, "summary": summary, "q": q, "stamp": stamp, "ticket": ticket,
        "meta": (f"{len(shown)} / {len(全員)} 件を表示" if 全員 else "会員データはありません"),
        "empty": ("条件に一致する会員はいません" if 全員 else "登録されている会員はいません"),
        "server_is_master": member_gate.会員はサーバーが正(),
        "gate_message": member_gate.断る文(),
        **_ガチャの画面(設定),
    })


@owner_required
@require_POST
def reward_gacha_save(request):
    """月別ガチャ特典設定の保存（旧アプリの saveRewardGachaConfig）。会員の表ではないので切り替えを待たない。"""
    入力 = _ガチャの入力(request)
    まずい = _ガチャを確かめる(入力)
    if まずい:
        messages.error(request, まずい)
        return reward_list(request, gacha_input=入力)
    答 = writes.ガチャ設定を保存する({"config": 入力})
    if 答.get("status") == "ok":
        messages.success(request, "月別ガチャ特典設定を保存しました")
    else:
        messages.error(request, "ガチャ特典設定の保存に失敗しました: " + (答.get("message") or "不明なエラー"))
    return redirect("manage:reward_list")


def _カード別(user, rewards):
    """旧アプリの buildStampCardDetails（詳細モーダルの「カード別スタンプ状況」）。"""
    いまの番号 = max(1, _数(user.get("stampCardNum"), 1) or 1)
    いまの数 = max(0, _数(user.get("stampCount")))
    枚数 = max(いまの番号, max((r["cardNum"] for r in rewards), default=0))
    出 = []
    for n in range(1, 枚数 + 1):
        最古 = timezone.make_aware(datetime(1970, 1, 1))
        分 = sorted((r for r in rewards if r["cardNum"] == n), key=lambda r: _時刻に(r["earnedDate"]) or 最古, reverse=True)
        いま = n == いまの番号
        未発行 = いま and いまの数 >= スタンプの上限 and not 分
        数 = _特典の数(分)
        if 未発行:
            状態 = {"label": "達成済み / 特典未発行", "badge": "badge-orange"}
        elif 数["unused"]:
            状態 = {"label": "未受取特典あり", "badge": "badge-green"}
        elif 数["expired"]:
            状態 = {"label": "期限切れ特典あり", "badge": "badge-red"}
        elif 数["used"]:
            状態 = {"label": "受取済み", "badge": "badge-blue"}
        elif いま and いまの数 > 0:
            状態 = {"label": "収集中", "badge": "badge-gray"}
        else:
            状態 = {"label": "履歴なし", "badge": "badge-gray"}
        出.append({
            "num": n, "current": いま,
            "stamp": いまの数 if いま else (スタンプの上限 if (分 or n < いまの番号) else 0),
            "last_stamp": _表示日(user.get("lastStampDate")) if いま else "ー",
            "achieved": _表示日(user.get("stampAchievedDate") if いま else (分[0]["earnedDate"] if 分 else "")),
            "pending": 未発行, "status": 状態, "counts": 数,
            "rewards": [{**r, "receipt": _受け取り状況(r), "expiry": _表示日(r["expiryDate"]),
                         "used_at": _表示日(r["usedAt"])} for r in 分],
        })
    return 出


def _特典の入力(request):
    """編集画面の特典の一覧を、旧アプリの saveRewardStatusEdit と同じ形に。"""
    出 = []
    for i in request.POST.getlist("reward"):
        出.append({
            "id": (request.POST.get(f"r_id_{i}") or "").strip() or f"reward-{timezone.now().timestamp():.0f}-{i}",
            "rewardName": (request.POST.get(f"r_name_{i}") or "").strip() or 特典の既定名,
            "rewardNote": (request.POST.get(f"r_note_{i}") or "").strip(),
            "cardNum": max(1, _数(request.POST.get(f"r_card_{i}"), 1) or 1),
            "earnedDate": _その日の始まり(request.POST.get(f"r_earned_{i}")),
            "expiryDate": _その日の終わり(request.POST.get(f"r_expiry_{i}")),
            "used": request.POST.get(f"r_used_{i}") == "on",
            "usedAt": _時刻の字(_時刻に(request.POST.get(f"r_used_at_{i}"))),
        })
    return _特典の一覧を整える(出)


def _特典の一覧を整える(一覧):
    """GAS の sanitizeRewardList_ と同じ。

    獲得日が空なら今、有効期限が空なら獲得日の1か月後。受取日時があれば受取済み。
    同じ（カード・特典名・獲得日）は1つにまとめ、獲得日の新しい順。
    GAS はシートに書く前にこれをしていたので、サーバーでも同じ姿で残す。
    """
    表 = {}
    for r in 一覧:
        獲得 = _時刻に(r["earnedDate"]) or timezone.localtime(timezone.now())
        期限 = _時刻に(r["expiryDate"]) or _一か月後(獲得)
        受取 = _時刻に(r["usedAt"])
        整 = {**r, "earnedDate": _時刻の字(獲得), "expiryDate": _時刻の字(期限),
             "used": r["used"] or bool(受取), "usedAt": _時刻の字(受取)}
        鍵 = f"{整['cardNum']}|{整['rewardName']}|{整['earnedDate']}"
        今 = 表.get(鍵)
        if not 今 or (整["used"] and not 今["used"]) or (整["usedAt"] and not 今["usedAt"]):
            表[鍵] = 整
    return sorted(表.values(), key=lambda r: _時刻に(r["earnedDate"]), reverse=True)


def _一か月後(t):
    """JS の setMonth(+1) と同じ（月末で溢れたぶんは翌月へ流れる）。"""
    年, 月 = (t.year + 1, 1) if t.month == 12 else (t.year, t.month + 1)
    return t.replace(year=年, month=月, day=1) + timedelta(days=t.day - 1)


@owner_required
def reward_edit(request, member_id: str):
    """詳細（旧アプリの 🔎 会員のスタンプ・特典詳細）と編集（🎟️ スタンプ・特典状況の編集）。"""
    m = get_object_or_404(Member, pk=member_id, deleted=False)
    if request.method == "POST":
        if not member_gate.会員はサーバーが正():
            messages.error(request, member_gate.断る文())
            return redirect("manage:reward_list")
        d = {
            "memberId": m.member_id,
            "stampCount": max(0, min(スタンプの上限, _数(request.POST.get("stampCount")))),
            "stampCardNum": max(1, _数(request.POST.get("stampCardNum"), 1) or 1),
            "lastStampDate": (request.POST.get("lastStampDate") or "").strip(),
            "stampAchievedDate": _時刻の字(_時刻に(request.POST.get("stampAchievedDate"))),
            "rewards": _特典の入力(request),
        }
        答 = admin_member.特典を書き換える(d)
        if 答.get("status") == "ok":
            messages.success(request, "スタンプ・特典状況を更新しました")
            return redirect("manage:reward_list")
        messages.error(request, 答.get("message") or "更新に失敗しました")

    状態 = admin_member._特典の状態(m)
    user = {**状態, "memberId": m.member_id, "name": m.name, "phone": m.phone}
    rewards = _特典一覧(user)
    数 = _特典の数(rewards)
    return render(request, "manage/reward_form.html", {
        "member": m, "server_is_master": member_gate.会員はサーバーが正(), "gate_message": member_gate.断る文(),
        "values": {
            "stampCount": 状態["stampCount"], "stampCardNum": 状態["stampCardNum"],
            "lastStampDate": _入力用の日(m.last_stamp_at), "stampAchievedDate": _入力用の日時(m.stamp_achieved_at),
        },
        "detail_summary": [("現在カード", f"{状態['stampCardNum']}枚目"), ("現在スタンプ", f"{状態['stampCount']} / {スタンプの上限}"),
                           ("未受取特典", 数["unused"]), ("受取済み特典", 数["used"])],
        "cards": _カード別(user, rewards),
        "entries": [{**r, "receipt": _受け取り状況(r), "expiry": _表示日(r["expiryDate"]), "used_at": _表示日(r["usedAt"]),
                     "earned_input": _入力用の日(r["earnedDate"]), "expiry_input": _入力用の日(r["expiryDate"]),
                     "used_at_input": _入力用の日時(r["usedAt"])} for r in rewards],
        "default_name": 特典の既定名,
    })
