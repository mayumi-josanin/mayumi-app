"""お知らせ管理。**中身は旧管理アプリ（#page-notice-management）と同じ。**

お客様アプリの「お知らせ一覧」に並ぶものを、NEWS・カレンダー・ショップ・ホームの
4つの表をまたいで、配信日時の新しい順に1つの表で見る画面。

ここで変えるのは「お知らせ一覧に出すか」だけ。元データ（本文や画像）は触らない。
書き込みは gasapi/admin_*.py の 一覧掲載を変える() / 一覧から外す()（GAS の
handleUpdateNoticeVisibility / handleDeleteNoticeListing の転送先）を使う。

並べ方・切り捨てる日付・カレンダーの休診と往診を外す規則は、
お客様アプリ（app.js の buildNoticeItems）と旧管理アプリの
buildAdminNoticeOverviewItems を写したもの。**ここだけ変えると、
管理画面とお客様アプリで並びが食い違う。**
"""

import datetime
import re

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import CalendarEvent, Category, Menu, News, Product
from apps.gasapi import admin_calendar, admin_menu, admin_news, admin_product

from .permissions import owner_required

# お知らせ一覧に出すのはこの日以降のもの（app.js の NOTICE_FEED_START_DATE と同じ）。
# 昔の記事まで並ぶと、お客様が新しい案内を見つけにくくなるため。
FEED_START_DATE = datetime.date(2026, 4, 1)

# 配信元ごとの、表の名前（GAS のシート名）・表示名・書き込み先・元データの編集画面
SOURCES = {
    "blog": {"sheet": "BLOG", "label": "NEWS", "api": admin_news, "edit": "manage:news_edit", "name": "NEWS記事"},
    "calendar": {"sheet": "CALENDAR", "label": "カレンダー", "api": admin_calendar, "edit": "manage:calendar_edit", "name": "カレンダー項目"},
    "product": {"sheet": "PRODUCTS", "label": "ショップ", "api": admin_product, "edit": "manage:product_edit", "name": "商品"},
    "menu": {"sheet": "MENUS", "label": "ホーム", "api": admin_menu, "edit": "manage:menu_edit", "name": "メニュー"},
}

# 「カテゴリ」欄に出す名前。カテゴリ管理に種別「通知」の名前があればそれを使う
# （旧管理アプリの getAdminManagedNoticeCategoryMap と同じ当て方）。
_CATEGORY_DEFAULTS = [
    ("blog", "NEWS", ["news", "ニュース"]),
    ("calendar", "カレンダー", ["カレンダー", "calendar"]),
    ("product", "ショップ", ["ショップ", "shop"]),
    ("menu", "ホーム", ["ホーム", "home"]),
]


def _照合の鍵(値: str) -> str:
    """カテゴリ名の照合用。小文字にし、ひらがなをカタカナに、空白と記号を落とす。"""
    文 = (値 or "").strip().lower()
    文 = "".join(chr(ord(ch) + 0x60) if "ぁ" <= ch <= "ん" else ch for ch in 文)
    return re.sub(r"[\s　_\-/]", "", 文)


def _配信元のカテゴリ名() -> dict:
    """配信元 → カテゴリ欄に出す名前。

    種別「通知」のカテゴリを、別名（news/ニュース など）で当てる。当たらなかった
    配信元には、余った「通知」のカテゴリを順に割り当て、それも無ければ既定の名前。
    """
    名たち, 見た = [], set()
    for c in Category.objects.filter(kind="通知").order_by("sheet_row"):
        名 = (c.name or "").strip()
        if 名 and 名 not in 見た:
            見た.add(名)
            名たち.append(名)
    使った = [False] * len(名たち)
    決めた = {}
    for kind, 既定, 別名 in _CATEGORY_DEFAULTS:
        鍵 = {_照合の鍵(x) for x in 別名 + [既定]}
        for i, 名 in enumerate(名たち):
            if not 使った[i] and _照合の鍵(名) in 鍵:
                決めた[kind] = 名
                使った[i] = True
                break
    余り = [名 for i, 名 in enumerate(名たち) if not 使った[i]]
    for kind, 既定, _ in _CATEGORY_DEFAULTS:
        if kind not in 決めた and 余り:
            決めた[kind] = 余り.pop(0)
        決めた.setdefault(kind, 既定)
    return 決めた


def _要約(text: str, 上限: int = 180) -> str:
    """一覧に出す本文。太字・下線の印と本文中の画像行を外し、1行にまとめて切り詰める。"""
    文 = re.sub(r"<\s*/?\s*(strong|b|u)\s*>", "", text or "", flags=re.I)
    文 = re.sub(r"(^|\n)\s*📷\s*https?://\S+", r"\1", 文)
    文 = re.sub(r"\s+", " ", 文).strip()
    if not 文:
        return "本文なし"
    return 文[:上限] + "…" if len(文) > 上限 else 文


def _日付を日時に(d):
    """投稿日などの「日付だけ」を、その日の0時（日本時間）にそろえる。"""
    if not d:
        return None
    return timezone.make_aware(datetime.datetime(d.year, d.month, d.day))


def _配信の日時(掲載日時, 更新日時, 元の日付=None, *, 並びに使う=False, 表示に使う=False, 期限に使う=False):
    """配信日時の決め方（旧管理アプリの buildAdminNoticePublishMeta と同じ）。

    お知らせ一覧に載せた日時 → 更新日時 → （配信元によっては）元の日付 の順。
    元の日付を使うかは配信元ごとに違う（NEWS・ホームは使う、カレンダーは期限の判定だけ、ショップは使わない）。
    返すのは (並べ替え用, 期限の判定用, 表示用の文字)。
    """
    元 = _日付を日時に(元の日付)
    並び = 掲載日時 or 更新日時 or (元 if 並びに使う else None)
    表示 = 掲載日時 or 更新日時 or (元 if 表示に使う else None)
    期限 = 掲載日時 or 更新日時 or (元 if 期限に使う else None)
    return 並び, 期限, (timezone.localtime(表示).strftime("%Y/%m/%d") if 表示 else "ー")


def _休診か(c: CalendarEvent) -> bool:
    """区分が入っていればそれに従う。空のときだけ題と説明から当てる（app.js と同じ規則）。"""
    区分 = (c.category or "").strip()
    if 区分:
        return bool(re.search(r"休診|休み|休業", 区分))
    return bool(re.search(r"休診|休み|休業", f"{c.title or ''} {c.detail or ''}"))


def _往診か(c: CalendarEvent) -> bool:
    区分 = (c.category or "").strip()
    if 区分:
        return "往診" in 区分
    return "往診" in (c.title or "")


def _一件(kind, obj, 日時, category, title, body, weight):
    並び, 期限, 表示 = 日時
    return {
        "kind": kind, "row": obj.sheet_row, "category": category, "source_label": SOURCES[kind]["label"],
        "status": "公開" if obj.notice_listed else "非公開", "title": title or "", "body": body or "",
        "summary": _要約(body), "date_label": 表示, "ts": 並び, "visible_ts": 期限,
        "sort_order": obj.sort_order or 0, "weight": weight, "edit_url": SOURCES[kind]["edit"],
    }


def _一覧を組む():
    """4つの表から、お知らせ一覧に出るものを集めて、配信日時の新しい順に並べる。"""
    分類 = _配信元のカテゴリ名()
    出 = []

    # NEWS: 公開中で、一覧から削除されていないもの。日時が無ければ投稿日を使う。
    記事 = list(News.objects.filter(deleted=False, published=True, notice_delisted_at__isnull=True).order_by("sheet_row"))
    for i, n in enumerate(記事):
        日時 = _配信の日時(n.notice_listed_at, n.updated_at, n.posted_on, 並びに使う=True, 表示に使う=True, 期限に使う=True)
        出.append(_一件("blog", n, 日時, 分類["blog"], n.title, n.body, len(記事) - i))

    # カレンダー: 休診と往診はお知らせ一覧に並べない。イベントの日付は期限の判定にだけ使う。
    予定 = [c for c in CalendarEvent.objects.filter(deleted=False, published=True, notice_delisted_at__isnull=True).order_by("sheet_row")
          if not _休診か(c) and not _往診か(c)]
    for i, c in enumerate(予定):
        日時 = _配信の日時(c.notice_listed_at, c.updated_at, c.event_on, 期限に使う=True)
        出.append(_一件("calendar", c, 日時, 分類["calendar"], c.title, c.detail, i))

    # ショップ: 元の日付は無い。
    商品 = list(Product.objects.filter(deleted=False, published=True, notice_delisted_at__isnull=True).order_by("sheet_row"))
    for i, p in enumerate(商品):
        日時 = _配信の日時(p.notice_listed_at, p.updated_at)
        出.append(_一件("product", p, 日時, 分類["product"], p.name, p.description, len(商品) - i))

    # ホーム（メニュー）: 日時が無ければ登録日を使う。
    献立 = list(Menu.objects.filter(deleted=False, published=True, notice_delisted_at__isnull=True).order_by("sheet_row"))
    for i, m in enumerate(献立):
        日時 = _配信の日時(m.notice_listed_at, m.updated_at, m.registered_on, 並びに使う=True, 表示に使う=True, 期限に使う=True)
        it = _一件("menu", m, 日時, 分類["menu"], m.name, m.summary, i)
        it["sort_order"] = m.sort_key or 0  # メニューの表示順は sort_key に入っている
        出.append(it)

    # 2026-04-01 より前のものは出さない（日時が分からないものは残す）。
    始まり = _日付を日時に(FEED_START_DATE)
    出 = [it for it in 出 if not (it["visible_ts"] or it["ts"]) or (it["visible_ts"] or it["ts"]) >= 始まり]
    # 新しい順 → 表示順の大きい順 → 配信元の中の重み
    底 = _日付を日時に(datetime.date(1970, 1, 1))
    出.sort(key=lambda it: (it["ts"] or 底, it["sort_order"], it["weight"]), reverse=True)
    return 出


def _検索用(値: str) -> str:
    """検索の照合用。小文字にし、空白を落とし、長音や全角ハイフンを「-」にそろえる。"""
    return re.sub(r"[‐－―ー]", "-", re.sub(r"\s+", "", (値 or "").lower()))


@owner_required
def notice_list(request):
    q = (request.GET.get("q") or "").strip()
    source = request.GET.get("source", "all")
    if source not in SOURCES:
        source = "all"
    all_items = _一覧を組む()
    shown = []
    鍵 = _検索用(q)
    for it in all_items:
        if source != "all" and it["kind"] != source:
            continue
        if 鍵 and 鍵 not in _検索用(" ".join([it["date_label"], it["category"], it["source_label"], it["status"], it["title"], it["body"]])):
            continue
        shown.append(it)
    summary = [
        ("配信総数", len(all_items)),
        ("NEWS", sum(1 for it in all_items if it["kind"] == "blog")),
        ("カレンダー", sum(1 for it in all_items if it["kind"] == "calendar")),
        ("ショップ", sum(1 for it in all_items if it["kind"] == "product")),
        ("ホーム", sum(1 for it in all_items if it["kind"] == "menu")),
        ("現在表示中", len(shown)),
    ]
    return render(request, "manage/notice_list.html", {
        "items": shown, "summary": summary, "q": q, "source": source,
        "sources": [(k, v["label"]) for k, v in SOURCES.items()],
        "meta": (f"{len(shown)} / {len(all_items)} 件を表示" if all_items else "配信中のお知らせはありません"),
        "empty": ("管理対象のお知らせはありません" if not all_items else "条件に一致する配信内容はありません"),
    })


def _宛先(request):
    """POST の kind / row を確かめて、配信元の定義と行番号を返す。おかしければ None。"""
    kind = request.POST.get("kind", "")
    try:
        row = int(request.POST.get("row") or 0)
    except ValueError:
        row = 0
    if kind not in SOURCES or row <= 1:
        return None, 0
    return SOURCES[kind], row


@owner_required
@require_POST
def notice_visibility(request):
    """「お知らせ一覧」欄の 公開/非公開 を切り替える（updateNoticeVisibility）。"""
    元, row = _宛先(request)
    if not 元:
        messages.error(request, "公開状態の更新対象が見つかりません")
        return redirect(request.POST.get("next") or "manage:notice_list")
    status = "非公開" if (request.POST.get("status") or "").strip() == "非公開" else "公開"
    答 = 元["api"].一覧掲載を変える({"sheet": 元["sheet"], "rowIdx": row, "status": status})
    if 答.get("status") == "ok":
        messages.success(request, "公開状態を更新しました")
    else:
        messages.error(request, "公開状態の更新に失敗しました: " + (答.get("message") or "不明なエラー"))
    return redirect(request.POST.get("next") or "manage:notice_list")


@owner_required
@require_POST
def notice_remove(request):
    """「一覧から削除」（deleteNoticeListing）。元データは残し、お知らせ一覧から下ろすだけ。

    ただしショップの商品一覧・ホームのメニュー一覧・カレンダーからも消える
    （お客様アプリはどれも「一覧から削除されていないこと」を見ているため）。
    この画面には二度と出てこないので、戻すときは元データの管理画面から。
    """
    元, row = _宛先(request)
    if not 元:
        messages.error(request, "削除対象が見つかりません")
        return redirect(request.POST.get("next") or "manage:notice_list")
    答 = 元["api"].一覧から外す({"sheet": 元["sheet"], "rowIdx": row})
    if 答.get("status") == "ok":
        messages.success(request, "お知らせ一覧から完全に削除しました")
    else:
        messages.error(request, "削除に失敗しました: " + (答.get("message") or "不明なエラー"))
    return redirect(request.POST.get("next") or "manage:notice_list")
