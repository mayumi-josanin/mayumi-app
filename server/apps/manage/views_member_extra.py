"""重複会員候補（統合）と、バックアップ / ゴミ箱。
**中身は旧管理アプリ（admin/index.html 会員管理ページの下半分: 🧩 重複会員候補・🛟 バックアップ / ゴミ箱）と同じ。**

重複候補の判定は GAS の buildDuplicateUsersFromRows_（旧アプリの buildDuplicateUsersFromAdminUsers）と同じ:
電話番号 → 氏名＋生年月日 → 氏名 の順に組を作り、同じ顔ぶれは1組にまとめて判定理由を並べる。
組の中はいちばん新しい会員が先頭で、そこへ他を統合する（「最新の会員IDへ統合」）。
「安全な候補を一括統合」は電話番号一致と氏名＋生年月日一致だけをつないだ組を対象にし、
氏名だけ一致は個別確認のまま残す（旧アプリの buildSafeDuplicateMergeComponentsFromUsers）。

統合は gasapi/admin_member.会員を統合する（注文の付け替えまで行う）。ゴミ箱の読み書きは gasapi/trash.py。

## 会員の表はまだスプレッドシートが正（9/19 まで）

統合と、会員の「復元」「完全削除」はどちらも会員の表への書き込みなので、
member_gate の印が立つまで断る。会員以外（NEWS・ショップ・カレンダー・ホーム・Push通知）は断らない。

## バックアップの状況はシステム管理にある

旧アプリはこの枠にバックアップ状況も出していたが、Django ではシステム管理に置いてある。
ここでは案内のリンクだけにする（同じものを2か所に置くと片方が古くなる）。
"""

import re

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils.dateparse import parse_datetime
from django.views.decorators.http import require_POST

from apps.gasapi import admin_member, trash

from . import member_gate
from .permissions import owner_required
from .views_system import _氏名を寄せる, _生年月日を寄せる, _電話を寄せる

# 判定理由の文言（GAS と同じ）
電話一致 = "電話番号が一致"
氏名生年月日一致 = "氏名と生年月日が一致"
氏名一致 = "氏名が一致"


# ---- 重複会員候補 ----

def _時刻の数(v):
    """並べ替え用。読めなければ 0（GAS の parseLooseDateToTimestamp_ と同じ扱い）。"""
    d = parse_datetime(str(v or "").strip().replace("/", "-")) if v else None
    return d.timestamp() if d else 0


def _表示の日付(v):
    """旧アプリの formatDisplayDate。YYYY/MM/DD、無ければ「ー」。"""
    s = str(v or "").strip()
    if not s:
        return "ー"
    d = parse_datetime(s.replace("/", "-").replace(" ", "T"))
    if d:
        return d.strftime("%Y/%m/%d")
    m = re.match(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})", s)
    return f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d}" if m else s


def _候補の1人(u):
    """組に出す項目（GAS の getDuplicateUsers が返す形と同じ）。"""
    return {
        "memberId": u.get("memberId") or "",
        "name": u.get("name") or "",
        "phone": u.get("phone") or "",
        "birthday": u.get("birthday") or "",
        "updatedAt": u.get("timestamp") or "",
        "updated_label": _表示の日付(u.get("timestamp")),
        "stampCount": u.get("stampCount") or 0,
        "orderCount": u.get("orderCount") or 0,
        "pendingOrderCount": u.get("pendingOrderCount") or 0,
        "orderTotal": u.get("orderTotal") or 0,
        "deviceCount": u.get("deviceCount") or 0,
    }


def _寄せて分ける(会員):
    """電話 / 氏名＋生年月日 / 氏名 の3つの辞書に分ける。"""
    電話, 氏名生年月日, 氏名 = {}, {}, {}
    for u in 会員:
        p, n, b = _電話を寄せる(u.get("phone")), _氏名を寄せる(u.get("name")), _生年月日を寄せる(u.get("birthday"))
        if p:
            電話.setdefault(p, []).append(u)
        if n and b:
            氏名生年月日.setdefault(n + "|" + b, []).append(u)
        if n:
            氏名.setdefault(n, []).append(u)
    return 電話, 氏名生年月日, 氏名


def _新しい順(顔ぶれ):
    return sorted(顔ぶれ, key=lambda u: _時刻の数(u.get("timestamp")), reverse=True)


def 重複候補(会員):
    """GAS の buildDuplicateUsersFromRows_ と同じ。組ごとに判定理由・候補（新しい順）。"""
    組, 見た = [], {}
    電話, 氏名生年月日, 氏名 = _寄せて分ける(会員)
    for 理由, 表 in ((電話一致, 電話), (氏名生年月日一致, 氏名生年月日), (氏名一致, 氏名)):
        for 顔ぶれ in 表.values():
            if len(顔ぶれ) < 2:
                continue
            鍵 = "|".join(sorted(str(u.get("memberId") or "") for u in 顔ぶれ))
            if 鍵 in 見た:
                if 理由 not in 見た[鍵]["reasons"]:
                    見た[鍵]["reasons"].append(理由)
                continue
            並び = [_候補の1人(u) for u in _新しい順(顔ぶれ)]
            g = {"key": 鍵, "reasons": [理由], "users": 並び,
                 "target": 並び[0], "sources": [x["memberId"] for x in 並び[1:]]}
            見た[鍵] = g
            組.append(g)
    return 組


def 安全な組(会員):
    """旧アプリの buildSafeDuplicateMergeComponentsFromUsers。

    電話番号一致・氏名＋生年月日一致でつながる会員をひとまとまりにし、新しい順に並べる。
    氏名だけ一致はつながない（別人のことがある）。
    """
    会員別 = {str(u.get("memberId") or "").strip(): u for u in 会員 if str(u.get("memberId") or "").strip()}
    隣 = {}
    電話, 氏名生年月日, _ = _寄せて分ける(会員)
    for 表 in (電話, 氏名生年月日):
        for 顔ぶれ in 表.values():
            ids = [str(u.get("memberId") or "").strip() for u in 顔ぶれ]
            ids = [x for x in ids if x]
            if len(ids) < 2:
                continue
            for a in ids:
                隣.setdefault(a, set()).update(x for x in ids if x != a)
    見た, 組 = set(), []
    for 始まり in 隣:
        if 始まり in 見た:
            continue
        山, まとまり = [始まり], []
        見た.add(始まり)
        while 山:
            mid = 山.pop()
            if mid in 会員別:
                まとまり.append(会員別[mid])
            for 次 in 隣.get(mid, ()):
                if 次 not in 見た:
                    見た.add(次)
                    山.append(次)
        if len(まとまり) > 1:
            組.append(_新しい順(まとまり))
    return 組


@owner_required
def duplicate_list(request):
    会員 = admin_member.一覧()["users"]
    組 = 重複候補(会員)
    return render(request, "manage/duplicates.html", {
        "groups": 組,
        "safe_count": len(安全な組(会員)),
        "member_writable": member_gate.会員はサーバーが正(),
    })


@owner_required
@require_POST
def duplicate_merge(request):
    """1組の統合（target + sources）か、安全な候補の一括統合（mode=safe）。"""
    if not member_gate.会員はサーバーが正():
        messages.error(request, member_gate.断る文())
        return redirect("manage:duplicate_list")

    if request.POST.get("mode") == "safe":
        # 画面を開いてから会員が変わっていることがあるので、いま時点で組み直す
        組 = 安全な組(admin_member.一覧()["users"])
        if not 組:
            messages.info(request, "一括統合できる安全な候補はありませんでした")
            return redirect("manage:duplicate_list")
        組数 = 件数 = 0
        for まとまり in 組:
            先 = str(まとまり[0].get("memberId") or "").strip()
            元 = [str(u.get("memberId") or "").strip() for u in まとまり[1:]]
            元 = [x for x in 元 if x]
            if not 先 or not 元:
                continue
            答 = admin_member.会員を統合する({"targetMemberId": 先, "sourceMemberIds": 元})
            if 答.get("status") != "ok":
                messages.error(request, "一括統合に失敗しました: " + (答.get("message") or "不明なエラー"))
                return redirect("manage:duplicate_list")
            組数 += 1
            件数 += len(元)
        if 組数:
            messages.success(request, f"安全な重複候補を {組数}組 ({件数}件) 統合しました")
        return redirect("manage:duplicate_list")

    先 = (request.POST.get("targetMemberId") or "").strip()
    元 = [x.strip() for x in request.POST.getlist("sourceMemberIds") if x.strip()]
    if not 先 or not 元:
        messages.error(request, "統合対象が不足しています")
        return redirect("manage:duplicate_list")
    答 = admin_member.会員を統合する({"targetMemberId": 先, "sourceMemberIds": 元})
    if 答.get("status") == "ok":
        messages.success(request, "会員を統合しました")
    else:
        messages.error(request, "統合に失敗しました" + (f": {答['message']}" if 答.get("message") else ""))
    return redirect("manage:duplicate_list")


# ---- バックアップ / ゴミ箱 ----

def _絞り込みの字(v):
    """旧アプリの normalizeFilterText。小文字にし、空白を除き、長音や全角の横棒を「-」にそろえる。"""
    return re.sub(r"[‐－―ー]", "-", re.sub(r"\s+", "", str(v or "").lower()))


def _項目の鍵(item):
    """行を指す印。会員は会員ID、それ以外はシートの行番号。"""
    return f"{item['sheet']}:{item['memberId'] or item['rowIdx']}"


@owner_required
def trash_list(request):
    q = (request.GET.get("q") or "").strip()
    source = request.GET.get("source") or "all"
    全部 = trash.一覧()["items"]
    語 = _絞り込みの字(q)
    shown = []
    for it in 全部:
        if source != "all" and (it.get("source") or "").strip() != source:
            continue
        if 語 and 語 not in _絞り込みの字(" ".join([it.get("title") or "", it.get("memberId") or "", it.get("reason") or "", it.get("source") or ""])):
            continue
        shown.append({**it, "key": _項目の鍵(it), "deleted_label": _表示の日付(it.get("deletedAt")),
                      "label": it.get("title") or it.get("memberId") or ""})
    return render(request, "manage/trash.html", {
        "items": shown, "q": q, "source": source,
        "sources": list(trash.種類の名前.values()),
        "empty": "条件に一致するゴミ箱項目はありません" if 全部 else "ゴミ箱は空です",
        "member_writable": member_gate.会員はサーバーが正(),
    })


def _選んだ項目(request):
    """POST の items（"SHEET:鍵"）を (sheet, 鍵) の並びに。形が違うものは捨てる。"""
    出 = []
    for v in request.POST.getlist("items"):
        sheet, _, 鍵 = (v or "").partition(":")
        sheet, 鍵 = sheet.strip().upper(), 鍵.strip()
        # 同じ行を2回選んでも1回だけ動かす
        if sheet and 鍵 and (sheet, 鍵) not in 出:
            出.append((sheet, 鍵))
    return 出


def _引数(sheet, 鍵):
    return {"sheet": sheet, "memberId": 鍵} if sheet == "USERS" else {"sheet": sheet, "rowIdx": 鍵}


def _ゴミ箱を操作する(request, 動かす, 選んでの文, 一括の文, 単独の文, 失敗の文):
    """復元と完全削除は流れが同じ。会員が混じっていて印が立っていなければ、会員だけ断る。"""
    選んだ = _選んだ項目(request)
    戻り先 = request.POST.get("next") or "manage:trash_list"
    if not 選んだ:
        messages.error(request, 選んでの文)
        return redirect(戻り先)
    会員は書ける = member_gate.会員はサーバーが正()
    できた = 0
    会員を断った = False
    for sheet, 鍵 in 選んだ:
        if sheet == "USERS" and not 会員は書ける:
            会員を断った = True
            continue
        if 動かす(_引数(sheet, 鍵)).get("status") == "ok":
            できた += 1
    if 会員を断った:
        messages.error(request, member_gate.断る文())
    if できた:
        messages.success(request, 一括の文 if request.POST.get("bulk") else 単独の文)
    elif not 会員を断った:
        messages.error(request, 失敗の文)
    return redirect(戻り先)


@owner_required
@require_POST
def trash_restore(request):
    return _ゴミ箱を操作する(request, trash.戻す, "復元する項目を選択してください", "一括復元しました", "復元しました", "復元に失敗しました")


@owner_required
@require_POST
def trash_hard_delete(request):
    return _ゴミ箱を操作する(request, trash.完全に消す, "完全削除する項目を選択してください", "一括完全削除しました", "完全削除しました", "完全削除に失敗しました")
