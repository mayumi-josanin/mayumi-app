"""カテゴリ管理。**中身は旧管理アプリ（#page-category-master カテゴリ管理）と同じ。**

読み書きは gasapi/admin_category.py の関数をそのまま使う。
旧管理アプリ（GAS 経由）と同じ決まり（名前が鍵・名前を変えても記事は変えない・
同じ名前は足せない）を、2か所に書かないため。

旧管理アプリの見せ方:
- 追加の分類は「NEWS（ブログ）」「メニュー一覧（メニュー）」「お知らせ一覧（通知）（通知）」の3つ。
  種別「お知らせ」は選べないが、残っている行は NEWS の表に出る（renderCategoryList と同じ）。
- 表は種別ごとに3つ（NEWS用 / メニュー一覧用 / お知らせ一覧（通知）用）。列は カテゴリ名・使用数・操作。
- 使用数は NEWS（getAdminBlogs）・メニュー（getAdminMenus）・お知らせ一覧（通知）の並びを
  カテゴリ名で数えたもの。**通知用カテゴリは記事に名前が書かれていない。**
  お知らせ一覧の各行が、出どころ（NEWS・カレンダー・ショップ・ホーム）から
  通知用カテゴリへ割り当てられる（getAdminManagedNoticeCategoryMap）ので、それを写す。
- 消すときは使用中でも消せる。「既存データのカテゴリ名はそのまま残る」と断ってから消す。
"""

from datetime import date, datetime

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.content.models import CalendarEvent, Menu, News, Product
from apps.gasapi import admin_category

from .permissions import owner_required

# 追加・編集で選べる分類（値と見せる文言）。旧管理アプリの select と同じ順。
分類の選択肢 = [
    ("ブログ", "NEWS"),
    ("メニュー", "メニュー一覧"),
    ("通知", "お知らせ一覧（通知）"),
]

# お知らせ一覧（通知）に載せるのはこの日以降の分だけ（旧管理アプリの ADMIN_NOTICE_FEED_START_DATE）。
お知らせ一覧の開始日 = date(2026, 4, 1)

# お知らせ一覧の出どころごとの、通知用カテゴリの既定の名前と別名（getAdminManagedNoticeCategoryMap の defaults）。
出どころの既定 = [
    ("blog", "NEWS", ["news", "ニュース"]),
    ("calendar", "カレンダー", ["カレンダー", "calendar"]),
    ("product", "ショップ", ["ショップ", "shop"]),
    ("menu", "ホーム", ["ホーム", "home"]),
]


@owner_required
def category_list(request):
    一覧 = admin_category.一覧()["categories"]
    使用数 = _使用数(一覧)
    表 = {"news": [], "menu": [], "notice": []}
    for 番号, c in enumerate(一覧):
        行 = {"no": 番号, "name": c["name"], "type": c["type"], "used": 使用数.get(c["name"], 0)}
        # 旧管理アプリと同じ分け方: メニュー・通知以外（ブログ・お知らせ）は NEWS の表。
        if c["type"] == "メニュー":
            表["menu"].append(行)
        elif c["type"] == "通知":
            表["notice"].append(行)
        else:
            表["news"].append(行)
    return render(request, "manage/category_list.html", {"tables": 表, "kinds": 分類の選択肢})


@owner_required
@require_POST
def category_add(request):
    答 = admin_category.足す(
        {"name": request.POST.get("name", ""), "categoryType": request.POST.get("kind", "")}
    )
    _知らせる(request, 答, "追加に失敗しました")
    return redirect("manage:category_list")


@owner_required
@require_POST
def category_update(request):
    答 = admin_category.書き換える(
        {
            "oldName": request.POST.get("old_name", ""),
            "newName": request.POST.get("name", ""),
            "categoryType": request.POST.get("kind", ""),
        }
    )
    _知らせる(request, 答, "更新に失敗しました")
    return redirect("manage:category_list")


@owner_required
@require_POST
def category_delete(request):
    # 旧管理アプリと同じく、使っている記事があっても消せる（記事側の名前は残る）。
    答 = admin_category.消す({"name": request.POST.get("name", "")})
    _知らせる(request, 答, "削除に失敗しました")
    return redirect("manage:category_list")


def _知らせる(request, 答, 失敗の頭):
    if 答.get("status") == "ok":
        messages.success(request, 答.get("message") or "保存しました。")
    else:
        messages.error(request, f"{失敗の頭}: {答.get('message') or '不明なエラー'}")


# ── 使用数 ─────────────────────────────────────────────


def _使用数(一覧):
    """旧管理アプリの renderCategoryList と同じ数え方。カテゴリ名 → 件数。

    NEWS（消していない・題のある記事）とメニュー（消していないもの）はカテゴリ名をそのまま数える。
    お知らせ一覧（通知）の各行は、出どころに割り当てた通知用カテゴリの名前で数える。
    """
    数 = {}

    def 足す(名):
        名 = (名 or "").strip()
        if 名:
            数[名] = 数.get(名, 0) + 1

    # getAdminBlogs は空のカテゴリを「お知らせ」として返す（admin_news._一件 と同じ）。
    for 名 in News.objects.filter(deleted=False).exclude(title="").values_list("category", flat=True):
        足す(名 or "お知らせ")
    for 名 in Menu.objects.filter(deleted=False).values_list("category", flat=True):
        足す(名)

    割り当て = _通知カテゴリの割り当て(一覧)
    for 出どころ, 件数 in _お知らせ一覧の件数().items():
        for _ in range(件数):
            足す(割り当て[出どころ])
    return 数


def _照合の鍵(値):
    """normalizeAdminCategoryLookupKey と同じ。小文字・ひらがな→カタカナ・空白と _ - / を除く。"""
    出 = []
    for ch in str(値 or "").strip().lower():
        if "ぁ" <= ch <= "ん":
            ch = chr(ord(ch) + 0x60)
        if ch in " 　_-/" or ch.isspace():
            continue
        出.append(ch)
    return "".join(出)


def _通知カテゴリの割り当て(一覧):
    """getAdminManagedNoticeCategoryMap と同じ。出どころ → 通知用カテゴリの名前。

    1. 既定の名前・別名に当てはまる通知用カテゴリを、その出どころに割り当てる。
    2. 余った通知用カテゴリを、まだ割り当てのない出どころへ順番に。
    3. それでも無ければ既定の名前（NEWS・カレンダー・ショップ・ホーム）。
    """
    候補 = [(c["name"], _照合の鍵(c["name"])) for c in 一覧 if c["type"] == "通知"]
    使った = set()
    割り当て = {}
    for 出どころ, 既定, 別名 in 出どころの既定:
        鍵 = {_照合の鍵(x) for x in 別名} | {_照合の鍵(既定)}
        for i, (名, k) in enumerate(候補):
            if i not in 使った and k in 鍵:
                割り当て[出どころ] = 名
                使った.add(i)
                break
    残り = [名 for i, (名, _) in enumerate(候補) if i not in 使った]
    for 出どころ, 既定, _ in 出どころの既定:
        if 出どころ not in 割り当て and 残り:
            割り当て[出どころ] = 残り.pop(0)
        if 出どころ not in 割り当て:
            割り当て[出どころ] = 既定
    return 割り当て


def _載る(掲載日時, 更新日時, 予備の日=None):
    """buildAdminNoticeOverviewItems の日付の絞り込み。日付が無ければ載る。あれば開始日以降だけ。"""
    値 = 掲載日時 or 更新日時 or 予備の日
    if not 値:
        return True
    if isinstance(値, datetime):
        値 = timezone.localdate(値)
    return 値 >= お知らせ一覧の開始日


def _休診か(c):
    """isAdminNoticeCalendarHolidayEvent と同じ。区分があればそれで見る。無ければ題と説明から。"""
    区分 = (c.category or "").strip()
    if 区分:
        return any(w in 区分 for w in ("休診", "休み", "休業"))
    文 = f"{c.title or ''} {c.detail or ''}"
    return any(w in 文 for w in ("休診", "休み", "休業"))


def _往診か(c):
    区分 = (c.category or "").strip()
    if 区分:
        return "往診" in 区分
    return "往診" in (c.title or "")


def _お知らせ一覧の件数():
    """お知らせ一覧（通知）に並ぶ件数を出どころごとに。buildAdminNoticeOverviewItems と同じ絞り込み。

    公開中で、お知らせ一覧から外していない（削除日時が空）ものだけ。
    カレンダーは休診と往診を並べない。
    """
    数 = {}
    生きている = dict(deleted=False, published=True, notice_delisted_at__isnull=True)

    数["blog"] = sum(
        _載る(n.notice_listed_at, n.updated_at, n.posted_on)
        for n in News.objects.filter(**生きている).exclude(title="")
    )
    数["calendar"] = sum(
        _載る(c.notice_listed_at, c.updated_at, c.event_on)
        for c in CalendarEvent.objects.filter(**生きている)
        if not _休診か(c) and not _往診か(c)
    )
    # 商品は登録日の予備が無い（旧管理アプリも updatedAt までしか見ない）。
    数["product"] = sum(
        _載る(p.notice_listed_at, p.updated_at) for p in Product.objects.filter(**生きている)
    )
    数["menu"] = sum(
        _載る(m.notice_listed_at, m.updated_at, m.registered_on)
        for m in Menu.objects.filter(**生きている)
    )
    return 数
