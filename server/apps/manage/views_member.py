"""会員管理。**中身は旧管理アプリ（admin/index.html の #page-users 会員一覧）と同じ。**

  集計カード   会員総数 / Push許可あり / 未対応注文あり / Push許可なし / 現在表示中
  絞り込み     検索（会員ID・氏名・フリガナ・電話番号・住所・メモ）/ 市区町村 / 年齢の下限上限 /
               並び替え6種 / Push状態3種 / 条件をクリア
  表の列       アイコン ID 登録/更新日時 氏名 電話番号 生年月日 年齢 住所 Push アンケート メモ 操作
               （＋ 表示する列を選べる。設計 docs/design/会員管理の画面を作り直す.md）
  操作         編集（旧アプリの「✏️ その場で」＝ 編集画面）/ 削除
  編集画面     編集できる項目・ボタン3つ（パスコードを設定・引き継ぎコードを発行・通知を止める）・
               見えるだけの項目（会員ID・登録日時・最終オンライン・統合先・登録経路の更新日時）

読み書きは gasapi/admin_member.py（GAS の転送先と同じ）。**新しい書き込み経路は作らない。**

## 会員の表は 9/19 までスプレッドシートが正

書き込みは member_gate.会員はサーバーが正() が False のあいだは断る。
読むのはよい（古いだけ）。切り替えの前にここから書くと、シートに無い記録ができて
正が2つになる。

## 旧アプリにあって、ここには無いもの

「オンライン」の緑の印は Firebase の在席情報（管理アプリだけが受け取る）で出していた。
サーバーには届かないので、端末の最終利用日時から「最近オンライン / 本日利用 / オフライン / 未同期」
だけを出す。
"""

import re
import unicodedata
from datetime import date, datetime

from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_datetime
from django.views.decorators.http import require_POST

from apps.gasapi import admin_member
from apps.members.models import Member

from . import member_gate
from .permissions import owner_required

# 並び替え（旧アプリの userSortSelect と同じ6種・同じ順）
並び替え = [
    ("registered_desc", "並び替え: 登録が新しい順"),
    ("registered_asc", "並び替え: 登録が古い順"),
    ("age_desc", "並び替え: 年齢が高い順"),
    ("age_asc", "並び替え: 年齢が低い順"),
    ("city_asc", "並び替え: 市区町村順"),
    ("kana_asc", "並び替え: 50音順（フリガナ）"),
]

# Push状態（旧アプリの userPushFilter と同じ3種）
PUSH状態 = [("all", "Push状態: 全て"), ("enabled", "Push許可あり"), ("disabled", "Push許可なし")]

市区町村が不明 = "__unknown__"

# 一覧の列。(鍵, 見出し, 既定で出すか)。既定で出す12列は旧アプリの表と同じ並び。
# それ以外は設計「一覧にも出す。表示する列を選べるように」で足した列。
# 「操作」は必ず出す（選べない）。
列一覧 = [
    ("avatar", "アイコン", True),
    ("memberId", "ID", True),
    ("timestamp", "登録/更新日時", True),
    ("name", "氏名", True),
    ("kana", "フリガナ", False),
    ("phone", "電話番号", True),
    ("birthday", "生年月日", True),
    ("age", "年齢", True),
    ("address", "住所", True),
    ("status", "ご状況", False),
    ("push", "Push", True),
    ("survey", "アンケート", True),
    ("stamp", "スタンプ", False),
    ("lastStampAt", "最終スタンプ取得日時", False),
    ("rewards", "未使用特典", False),
    ("role", "権限", False),
    ("bijiris", "ビジリス", False),
    ("lineUserId", "LINEユーザーID", False),
    ("registrationSource", "登録経路", False),
    ("lastOnline", "最終オンライン", False),
    ("mergedIntoId", "統合先会員ID", False),
    ("transferCode", "引き継ぎコード", False),
    ("memo", "メモ", True),
]
既定の列 = [k for k, _, on in 列一覧 if on]


# ---- 旧アプリの JS と同じ計算 ----

def _文(v):
    return "" if v is None else str(v)


def _日時を読む(v):
    """'2026-09-07T14:30:00+09:00' も '2026/9/7 14:30' も読む。読めなければ None。"""
    文 = _文(v).strip()
    if not 文:
        return None
    d = parse_datetime(文.replace("/", "-").replace(" ", "T", 1))
    if d is None:
        m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", 文)
        if not m:
            return None
        try:
            d = datetime(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            return None
    if timezone.is_naive(d):
        d = timezone.make_aware(d)
    return timezone.localtime(d)


def _日付の表示(v):
    """旧アプリの formatDisplayDate。無ければ「ー」。"""
    d = _日時を読む(v)
    return d.strftime("%Y/%m/%d") if d else ("ー" if not _文(v).strip() else _文(v))


def _日時の表示(v):
    """旧アプリの formatSurveyTimestamp / formatAdminOrderDateTime（'YYYY/MM/DD HH:MM'）。"""
    d = _日時を読む(v)
    return d.strftime("%Y/%m/%d %H:%M") if d else ""


def _生年月日を読む(v):
    """旧アプリの parseUserBirthday。'1988/4/22' も '1988-04-22' も '1988年4月22日' も読む。"""
    m = re.match(r"^(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})", _文(v).strip())
    if not m:
        return None
    try:
        return date(int(m[1]), int(m[2]), int(m[3]))  # 2月30日のような日付は ValueError
    except ValueError:
        return None


def _年齢(birthday, きょう=None):
    """旧アプリの getUserAge。生年月日が読めなければ None。0〜130 の範囲外も None。"""
    b = _生年月日を読む(birthday)
    if not b:
        return None
    きょう = きょう or timezone.localdate()
    age = きょう.year - b.year - ((きょう.month, きょう.day) < (b.month, b.day))
    return age if 0 <= age <= 130 else None


def _市区町村(address):
    """旧アプリの getUserCity。住所の表記ゆれを吸収して市区町村だけを取り出す。

    「厚木市…」「神奈川県厚木市…」を同じ扱いにし、郡がある場合は町村名を採る。
    「神」が欠けた「奈川県…」のような入力もあるため、都道府県の判定はゆるめにしている。
    """
    文 = re.sub(r"[\s　]+", "", _文(address).strip())
    if not 文:
        return ""
    文 = re.sub(r"^(東京都|北海道|(?:京都|大阪)府|.{2,3}県)", "", 文)
    郡 = re.match(r"^.{1,6}郡(.{1,6}?[町村])", 文)
    if 郡:
        return 郡[1]
    市 = re.match(r"^(.{1,8}?[市区町村])", 文)
    return 市[1] if 市 else ""


def _検索用(text):
    """旧アプリの normalizeFilterText。小文字・空白なし・長音や全角ハイフンは '-' に。"""
    return re.sub(r"[‐－―ー]", "-", re.sub(r"\s+", "", _文(text).lower()))


def _かなの鍵(u):
    """50音順の鍵。**フリガナだけで並べる。**（旧アプリの kanaKey）

    フリガナが無い方は末尾にまとめる。五十音のどこに置いても嘘になるので混ぜない。
    """
    return re.sub(r"[\s　]+", "", unicodedata.normalize("NFKC", _文(u.get("kana"))))


def _在席(u):
    """旧アプリの getAdminUserPresenceMeta。端末の最終利用日時から状態の札を出す。

    「オンライン」（Firebase の在席）はサーバーからは分からないので出さない。
    """
    最終 = None
    for s in (u.get("deviceSessions") or []):
        d = _日時を読む((s or {}).get("lastSeenAt")) if isinstance(s, dict) else None
        if d and (最終 is None or d > 最終):
            最終 = d
    if not 最終:
        d = _日時を読む(u.get("lastOnlineAt"))
        最終 = d
    if not 最終:
        return {"label": "未同期", "badge": "badge-gray", "lastSeen": "---"}
    分 = max(0, (timezone.now() - 最終).total_seconds() / 60)
    表示 = 最終.strftime("%Y/%m/%d %H:%M")
    if 分 <= 10:
        return {"label": "最近オンライン", "badge": "badge-green", "lastSeen": 表示}
    if 分 <= 1440:
        return {"label": "本日利用", "badge": "badge-blue", "lastSeen": 表示}
    return {"label": "オフライン", "badge": "badge-gray", "lastSeen": 表示}


def _アンケート(u):
    """旧アプリの buildSurveyHistoryMarkup と同じ行。回答とスタンプ付与は別に記録される。"""
    回答 = _日時の表示(u.get("surveyAnsweredAt"))
    付与 = _日時の表示(u.get("surveyStampGrantedAt"))
    保留 = _日時の表示(u.get("surveyStampPendingAt"))
    if not (回答 or 付与 or 保留):
        return None
    行 = [("回答 " + 回答) if 回答 else "回答 記録なし"]
    if 付与:
        行.append("スタンプ " + 付与)
    elif 保留:
        行.append(f"スタンプ 保留中（{保留}）")
    else:
        行.append("スタンプ 未付与")
    return 行


def _未使用特典の数(u):
    return sum(1 for r in (u.get("rewards") or []) if isinstance(r, dict) and not r.get("used"))


def _一件(u):
    """一覧の1行に出す形。旧アプリの renderUsers が行ごとに計算していたもの。"""
    画像 = _文(u.get("avatarUrl")).strip()
    return {
        "u": u,
        "avatar": 画像 if 画像.startswith(("http://", "https://", "data:", "/")) else "",
        "initial": (_文(u.get("name")).strip() or "👤")[:1],
        "timestamp": _日付の表示(u.get("timestamp")),
        "birthday": _日付の表示(u.get("birthday")) if _文(u.get("birthday")).strip() else "",
        "age": _年齢(u.get("birthday")),
        "city": _市区町村(u.get("address")),
        "presence": _在席(u),
        "deviceCount": max(int(u.get("deviceCount") or 0), len(u.get("deviceSessions") or [])),
        "survey": _アンケート(u),
        "lastOrder": _日付の表示(u.get("lastOrderAt")) if _文(u.get("lastOrderAt")).strip() else "---",
        "lastStampAt": _日時の表示(u.get("lastStampAt")),
        "lastOnline": _日時の表示(u.get("lastOnlineAt")),
        "unusedRewards": _未使用特典の数(u),
    }


def _並べる(users, mode):
    """旧アプリの sortFilteredUsers。生年月日・市区町村・フリガナが無い方は末尾にまとめる。"""
    出 = list(users)
    if mode == "age_desc":
        出.sort(key=lambda u: (u["age"] is None, -(u["age"] or 0)))
    elif mode == "age_asc":
        出.sort(key=lambda u: (u["age"] is None, u["age"] or 0))
    elif mode == "city_asc":
        出.sort(key=lambda u: (not u["city"], u["city"]))
    elif mode == "kana_asc":
        # 同じ読みの方が続くときは、並びが毎回変わらないようお名前で決める
        出.sort(key=lambda u: (not _かなの鍵(u["u"]), _かなの鍵(u["u"]), _文(u["u"].get("name"))))
    elif mode == "registered_asc":
        出.reverse()
    return 出


def _新しい順(users):
    """GAS の getAdminUsers は「最新を上に」返す。サーバー版は会員IDの昇順なので、ここで登録日時の新しい順に。"""
    return sorted(users, key=lambda u: (_文(u.get("timestamp")), _文(u.get("memberId"))), reverse=True)


def _全員():
    """admin_member.一覧()（GAS の getAdminUsers と同じ31項目）に、編集画面と選べる列が使う項目を足す。

    GAS のシート版はこれらを返さない（一覧() の項目は GAS と同じに保つ約束があり、
    試験でも項目の集合を固めている）。旧管理アプリの編集画面（admin/index.html の editUser）は
    この名前で読むので、**同じ名前で**ここだけで足す。
    パスワードハッシュ・ソルト・届け先そのものは足さない（設計で「画面に出さない」）。
    """
    追加 = {
        m.member_id: {
            "role": m.role or "", "bijirisRegistered": bool(m.bijiris_registered),
            "lineUserId": m.line_user_id or "", "lastOnlineAt": m.last_online_at.isoformat() if m.last_online_at else "",
            "mergedIntoId": m.merged_into_id or "", "transferCode": m.transfer_code or "",
        }
        for m in Member.objects.exclude(deleted=True).only(
            "member_id", "role", "bijiris_registered", "line_user_id", "last_online_at", "merged_into_id", "transfer_code")
    }
    出 = []
    for u in admin_member.一覧()["users"]:
        u.update(追加.get(u["memberId"], {}))
        出.append(u)
    return _新しい順(出)


@owner_required
def member_list(request):
    全員 = _全員()
    q = _検索用(request.GET.get("q") or "")
    city = request.GET.get("city") or "all"
    push = request.GET.get("push") or "all"
    sort = request.GET.get("sort") or "registered_desc"
    if sort not in dict(並び替え):
        sort = "registered_desc"
    age_min = (request.GET.get("age_min") or "").strip()
    age_max = (request.GET.get("age_max") or "").strip()
    下限 = int(age_min) if age_min.isdigit() else None
    上限 = int(age_max) if age_max.isdigit() else None
    cols = request.GET.getlist("col") or 既定の列
    if request.GET.get("cols_given") and not request.GET.getlist("col"):
        cols = []

    # 市区町村の候補（件数の多い順。同数なら名前順）。旧アプリの refreshUserCityFilterOptions
    件数 = {}
    不明 = 0
    for u in 全員:
        c = _市区町村(u.get("address"))
        if c:
            件数[c] = 件数.get(c, 0) + 1
        else:
            不明 += 1
    市区町村の候補 = sorted(件数.items(), key=lambda x: (-x[1], x[0]))

    表示 = []
    for u in 全員:
        if push == "enabled" and not u.get("pushEnabled"):
            continue
        if push == "disabled" and u.get("pushEnabled"):
            continue
        c = _市区町村(u.get("address"))
        if city != "all":
            if city == 市区町村が不明:
                if c:
                    continue
            elif c != city:
                continue
        if 下限 is not None or 上限 is not None:
            # 生年月日が未登録だと年齢を判定できないため、年齢で絞る間は除外する
            age = _年齢(u.get("birthday"))
            if age is None:
                continue
            if 下限 is not None and age < 下限:
                continue
            if 上限 is not None and age > 上限:
                continue
        if q:
            haystack = _検索用(" ".join(_文(u.get(k)) for k in ("memberId", "name", "kana", "phone", "birthday", "address", "memo")))
            if q not in haystack:
                continue
        表示.append(_一件(u))
    表示 = _並べる(表示, sort)

    summary = [
        ("会員総数", len(全員)),
        ("Push許可あり", sum(1 for u in 全員 if u.get("pushEnabled"))),
        ("未対応注文あり", sum(1 for u in 全員 if int(u.get("pendingOrderCount") or 0) > 0)),
        ("Push許可なし", sum(1 for u in 全員 if not u.get("pushEnabled"))),
        ("現在表示中", len(表示)),
    ]
    return render(request, "manage/member_list.html", {
        "items": 表示, "summary": summary, "q": request.GET.get("q") or "", "city": city, "push": push, "sort": sort,
        "age_min": age_min, "age_max": age_max, "sort_choices": 並び替え, "push_choices": PUSH状態,
        "cities": 市区町村の候補, "unknown_city": 不明, "unknown_key": 市区町村が不明,
        "columns": [(k, label, k in cols) for k, label, _ in 列一覧], "cols": cols,
        "meta": (f"{len(表示)} / {len(全員)} 件を表示" if 全員 else "会員データはありません"),
        "empty": ("条件に一致する会員はいません" if 全員 else "登録されている会員はいません"),
        "writable": member_gate.会員はサーバーが正(), "gate_message": member_gate.断る文(),
    })


# ---- 編集 ----

def _会員(member_id: str):
    """**会員番号で探す。**お名前では探さない。消した会員は編集しない。"""
    return get_object_or_404(Member, pk=member_id, deleted=False)


def _書けるか(request):
    """9/19 の切り替えまで、会員まわりの書き込みは断る。"""
    if member_gate.会員はサーバーが正():
        return True
    messages.error(request, member_gate.断る文())
    return False


def _値(m: Member) -> dict:
    """編集画面の入力欄に入れる値。GAS の getAdminUsers と同じ名前で持つ。"""
    return {
        "name": m.name or "", "kana": m.kana or "", "phone": m.phone or "",
        "birthday": m.birthday.strftime("%Y-%m-%d") if m.birthday else "",
        "address": m.address or "", "memo": m.memo or "", "status": m.status or "",
        "role": m.role or "", "bijiris": "登録済み" if m.bijiris_registered else "",
        "stampCount": m.stamp_count or 0, "stampCardNum": m.stamp_card_number or 0,
        "lastStampAt": timezone.localtime(m.last_stamp_at).strftime("%Y-%m-%dT%H:%M") if m.last_stamp_at else "",
        "registrationSource": m.registration_source or "",
        "registrationSourceDetail": m.registration_source_detail or "",
        "lineUserId": m.line_user_id or "",
    }


def _画面(request, m: Member, values):
    def 日時(d):
        return timezone.localtime(d).strftime("%Y/%m/%d %H:%M") if d else "—"

    u = next((x for x in admin_member.一覧()["users"] if x["memberId"] == m.member_id), None) or {}
    return render(request, "manage/member_form.html", {
        "member": m, "values": values,
        # 見えるだけの項目（触ると壊れるもの）
        "view": {
            "memberId": m.member_id, "timestamp": 日時(m.created_at), "lastOnline": 日時(m.last_online_at),
            "mergedIntoId": m.merged_into_id or "—", "registrationSourceUpdatedAt": 日時(m.registration_source_updated_at),
            "deviceCount": len(list(m.device_sessions or [])),
            # **届け先そのものは出さない。**長い英数字で、見ても意味がない。
            "push": "設定あり" if (_文(m.push_subscription).strip() or m.push_enabled) else "なし",
            "transferCode": m.transfer_code or "—", "transferCodeIssuedAt": 日時(m.transfer_code_issued_at),
            "lastStampDate": timezone.localtime(m.last_stamp_at).strftime("%Y/%m/%d") if m.last_stamp_at else "—",
            "stampAchievedAt": 日時(m.stamp_achieved_at),
            "orderCount": u.get("orderCount", 0), "pendingOrderCount": u.get("pendingOrderCount", 0),
            "unusedRewards": _未使用特典の数(u),
        },
        "writable": member_gate.会員はサーバーが正(), "gate_message": member_gate.断る文(),
    })


@owner_required
def member_edit(request, member_id: str):
    m = _会員(member_id)
    if request.method == "POST":
        if not _書けるか(request):
            return redirect("manage:member_list")
        P = request.POST
        if not (P.get("name") or "").strip():
            messages.error(request, "氏名を入力してください。")
            return _画面(request, m, P)
        # 旧アプリの saveUserEdit / saveUserInline が送っていた項目と同じ。
        # 電話番号の先頭の0・LINEユーザーIDの重複は 会員を書き換える が見る。
        答 = admin_member.会員を書き換える({
            "memberId": m.member_id,
            "name": P.get("name", ""), "kana": P.get("kana", ""), "phone": P.get("phone", ""),
            "birthday": P.get("birthday", ""), "address": P.get("address", ""), "memo": (P.get("memo") or "").strip(),
            "status": P.get("status", ""), "role": P.get("role", ""),
            "bijirisRegistered": P.get("bijiris") == "登録済み",
            "stampCount": P.get("stampCount") or 0, "stampCardNum": P.get("stampCardNum") or 0,
            "lastStampAt": P.get("lastStampAt", ""),
            "registrationSource": P.get("registrationSource", ""),
            "registrationSourceDetail": P.get("registrationSourceDetail", ""),
            "lineUserId": P.get("lineUserId", ""),
        })
        if 答.get("status") == "ok":
            messages.success(request, "会員情報を更新しました")
            return redirect("manage:member_list")
        messages.error(request, 答.get("message") or "更新に失敗しました")
        return _画面(request, m, P)
    return _画面(request, m, _値(m))


@owner_required
@require_POST
def member_passcode(request, member_id: str):
    """受付で「パスコードを忘れた」と言われたときに使う。**いまの値は見られない。**"""
    m = _会員(member_id)
    if not _書けるか(request):
        return redirect("manage:member_list")
    新 = (request.POST.get("passcode") or "").strip()
    if not re.match(r"^\d{4}$|^\d{6}$", 新):
        messages.error(request, "パスコードは数字4桁または6桁で入力してください。")
        return redirect("manage:member_edit", member_id=m.member_id)
    答 = admin_member.パスコードを設定する({"memberId": m.member_id, "passcode": 新})
    if 答.get("status") == "ok":
        messages.success(request, "パスコードを設定しました")
    else:
        messages.error(request, 答.get("message") or "パスコードを設定できませんでした")
    return redirect("manage:member_edit", member_id=m.member_id)


@owner_required
@require_POST
def member_transfer_code(request, member_id: str):
    """機種変更のときに使う8桁のコード。**手で入れさせない**（重複すると別の方の記録が復元される）。

    発行したコードと有効期限は画面に出す。お客様に伝えるため。
    """
    m = _会員(member_id)
    if not _書けるか(request):
        return redirect("manage:member_list")
    答 = admin_member.引き継ぎコードを出す({"memberId": m.member_id})
    if 答.get("status") == "ok":
        messages.success(request, f"引き継ぎコード: {答.get('transferCode', '')}　有効期限: {答.get('expiresAtLabel') or '1週間'}　お客様にお伝えください。")
    else:
        messages.error(request, 答.get("message") or "発行できませんでした")
    return redirect("manage:member_edit", member_id=m.member_id)


@owner_required
@require_POST
def member_stop_push(request, member_id: str):
    """通知を止める。**再開はお客様の端末から。**"""
    m = _会員(member_id)
    if not _書けるか(request):
        return redirect("manage:member_list")
    答 = admin_member.通知を止める({"memberId": m.member_id})
    if 答.get("status") == "ok":
        messages.success(request, "通知を止めました")
    else:
        messages.error(request, 答.get("message") or "止められませんでした")
    return redirect("manage:member_edit", member_id=m.member_id)


@owner_required
@require_POST
def member_delete(request, member_id: str):
    """印を付けるだけ（退会）。行は残る。旧アプリの deleteUser と同じく「ゴミ箱へ移動」。"""
    m = _会員(member_id)
    if not _書けるか(request):
        return redirect("manage:member_list")
    答 = admin_member.会員を消す({"memberId": m.member_id})
    if 答.get("status") == "ok":
        messages.success(request, "会員を削除しました")
    else:
        messages.error(request, "削除に失敗しました: " + (答.get("message") or "不明なエラー"))
    return redirect(request.POST.get("next") or "manage:member_list")
