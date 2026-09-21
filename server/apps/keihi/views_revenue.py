"""売上の記録を、日ごとにまとめて現金出納帳へ入れる。

院長の依頼（2026-09-21）:
    「売上の記録（メニュー）・売上の記録（商品）にその日の売上を記録していくので、
     経費管理に**日にちごとの売上の合計金額**を反映してほしい。
     例: 9/20 にメニュー合計100,000円・商品合計50,000円なら、9/20 の売上は150,000円」

決めごと（院長の判断 2026-09-21）:

    **ボタンを押して入れる。**売上を記録した時点で帳簿が勝手に変わると、
    院長が手で直した行が黙って書き換わる。押したときだけ入れる。

    **1日1行。**メニューと商品を分けない。帳簿が見やすく、税理士に出すときも扱いやすい。

    **下見 → 取り込み の2段。**押していきなり入れない。
    何日ぶんがいくらで入るかを見てから決められるようにする。
    レシートの「確かめてから出納帳へ」と同じ考え方（views_receipt.receipt_to_book）。

金額の式は**売上の記録・データ分析と同じもの**を使う（`analytics_calc.record_totals`）。
ここで別の式を作ると、同じ月の売上が画面によって違う額になり、
院長がどちらを信じてよいか分からなくなる。
"""

from django.contrib import messages
from django.db import transaction
from django.shortcuts import redirect, render

from apps.manage.analytics_calc import record_totals
from apps.manage.permissions import owner_required
from apps.records.models import RevenueRecord

from .models import Cashbook, CashbookEntry
from .views import _帳簿を用意する, _選んだ年月

# 取り込んだ行だと分かるようにする印。摘要の先頭に付ける。
# **院長が手で直した行と見分けるため。**印が付いた行だけを「すでに入っている」と数える。
印 = "売上"


def _その月の日ごとの売上(year: int, month: int) -> dict[int, dict]:
    """その月の、日ごとの売上をまとめる。

    返すもの: {日: {"menu": メニューの合計, "product": 商品の合計,
                    "total": 合計, "menu_count": 件数, "product_count": 件数}}
    """
    日ごと: dict[int, dict] = {}
    for r in RevenueRecord.objects.filter(deleted=False, recorded_on__year=year, recorded_on__month=month):
        if not r.recorded_on:
            continue
        t = record_totals(r)
        if t["qty"] <= 0:
            continue
        日 = 日ごと.setdefault(r.recorded_on.day, {
            "menu": 0, "product": 0, "total": 0, "menu_count": 0, "product_count": 0,
        })
        額 = int(round(t["totalAmount"]))
        if r.kind == RevenueRecord.MENU:
            日["menu"] += 額
            日["menu_count"] += 1
        else:
            日["product"] += 額
            日["product_count"] += 1
        日["total"] += 額
    # 0円の日は行を作らない（無料・サービスだけの日に「売上 0円」と載せても意味がない）
    return {日: v for 日, v in 日ごと.items() if v["total"] > 0}


def _すでに入っている行(book, month: int) -> dict[int, CashbookEntry]:
    """その帳簿に、取り込んだ売上の行がすでにあるか。日で引ける形にする。"""
    出来上がり = {}
    for e in book.entries.all():
        if e.day and (e.description or "").startswith(印) and e.income:
            出来上がり[e.day] = e
    return 出来上がり


def _売上のある月() -> list[tuple[int, int]]:
    """売上の記録がある年月を、古い順に。"""
    月 = set()
    for d in RevenueRecord.objects.filter(deleted=False).exclude(recorded_on=None).values_list(
            "recorded_on", flat=True):
        月.add((d.year, d.month))
    return sorted(月)


@owner_required
def revenue_all_preview(request):
    """**これまでの全部の月**の下見。月ごとの合計と、帳簿に入っているかを並べる。

    6か月ぶん80日を1つずつ入れるのは大変なので、まとめて入れられるようにする。
    ただし**ここでも押すまでは入れない。**月ごとの下見と同じ考え方。
    """
    行 = []
    for year, month in _売上のある月():
        売上 = _その月の日ごとの売上(year, month)
        if not 売上:
            continue
        book = Cashbook.objects.filter(year=year, month=month).first()
        既存 = _すでに入っている行(book, month) if book else {}
        合計 = sum(v["total"] for v in 売上.values())
        新しい = [日 for 日 in 売上 if 日 not in 既存]
        違う = [日 for 日, v in 売上.items() if 日 in 既存 and 既存[日].income != v["total"]]
        行.append({
            "year": year, "month": month,
            "days": len(売上), "total": 合計,
            "menu": sum(v["menu"] for v in 売上.values()),
            "product": sum(v["product"] for v in 売上.values()),
            "new_days": len(新しい), "differ_days": len(違う),
            "done": (not 新しい and not 違う),
        })
    return render(request, "keihi/revenue_all.html", {
        "rows": 行,
        "total": sum(r["total"] for r in 行),
        "new_total": sum(r["new_days"] for r in 行),
    })


@owner_required
def revenue_all_to_book(request):
    """選んだ月を、まとめて出納帳へ入れる。中身は月ごとの取り込みと同じ決まり。"""
    if request.method != "POST":
        return redirect("manage:keihi:revenue_all_preview")

    やり方 = request.POST.get("on_conflict") or "skip"
    選んだ月 = set()
    for x in request.POST.getlist("months"):
        年月 = str(x).split("-")
        if len(年月) == 2 and 年月[0].isdigit() and 年月[1].isdigit():
            選んだ月.add((int(年月[0]), int(年月[1])))

    入れた日, 直した日, 飛ばした日, 月数 = 0, 0, 0, 0
    for year, month in sorted(選んだ月):
        i, n, s = _1か月ぶん入れる(year, month, set(_その月の日ごとの売上(year, month)), やり方)
        入れた日 += i
        直した日 += n
        飛ばした日 += s
        if i or n:
            月数 += 1

    if 入れた日:
        messages.success(request, f"{月数} か月ぶん・{入れた日} 日ぶんを現金出納帳に入れました")
    if 直した日:
        messages.success(request, f"{直した日} 日ぶんを新しい金額に直しました")
    if 飛ばした日:
        messages.info(request, f"{飛ばした日} 日ぶんは、すでに同じ金額が入っているので触っていません")
    return redirect("manage:keihi:revenue_all_preview")


@owner_required
def revenue_preview(request):
    """下見。その月の日ごとの売上と、すでに帳簿に入っている額を並べて見せる。

    **売上をあとから直したときに気づけるように**、両方を出す。
    """
    year, month = _選んだ年月(request)
    book = _帳簿を用意する(year, month)
    売上 = _その月の日ごとの売上(year, month)
    既存 = _すでに入っている行(book, month)

    行 = []
    for 日 in sorted(売上):
        v = 売上[日]
        いま = 既存.get(日)
        行.append({
            "day": 日,
            "menu": v["menu"],
            "product": v["product"],
            "total": v["total"],
            "menu_count": v["menu_count"],
            "product_count": v["product_count"],
            "in_book": (いま.income if いま else None),
            "same": bool(いま and いま.income == v["total"]),
            "differs": bool(いま and いま.income != v["total"]),
        })

    return render(request, "keihi/revenue_preview.html", {
        "year": year, "month": month, "rows": 行,
        "total": sum(r["total"] for r in 行),
        "new_count": sum(1 for r in 行 if r["in_book"] is None),
        "same_count": sum(1 for r in 行 if r["same"]),
        "differ_count": sum(1 for r in 行 if r["differs"]),
    })


@owner_required
def revenue_to_book(request):
    """下見で選んだ日を、出納帳の行にする。

    **二度入れない。**同じ日の売上の行があるときは:
      ・金額が同じ … 飛ばす
      ・金額が違う … 選んだやり方で（直す／そのまま／別の行として足す）
    院長が手で直した行を、黙って書き換えないため。
    """
    if request.method != "POST":
        return redirect("manage:keihi:revenue_preview")

    year, month = _選んだ年月(request)
    やり方 = request.POST.get("on_conflict") or "skip"   # replace / skip / add
    選んだ日 = {int(d) for d in request.POST.getlist("days") if str(d).isdigit()}

    入れた, 直した, 飛ばした = _1か月ぶん入れる(year, month, 選んだ日, やり方)

    if 入れた:
        messages.success(request, f"{入れた} 日ぶんを現金出納帳に入れました")
    if 直した:
        messages.success(request, f"{直した} 日ぶんを新しい金額に直しました")
    if 飛ばした:
        messages.info(request, f"{飛ばした} 日ぶんは、すでに同じ金額が入っているので触っていません")
    return redirect(f"/manage/keihi/?year={year}&month={month}")


def _1か月ぶん入れる(year: int, month: int, 選んだ日: set, やり方: str) -> tuple[int, int, int]:
    """1か月ぶんを出納帳へ入れる。月ごとの取り込みと、まとめての取り込みが共に通る所。

    **同じ決まりを2か所に書かない。**片方だけ直すと、月ごととまとめてで動きが変わる。
    """
    book = _帳簿を用意する(year, month)
    売上 = _その月の日ごとの売上(year, month)
    既存 = _すでに入っている行(book, month)

    入れた, 直した, 飛ばした = 0, 0, 0
    with transaction.atomic():
        for 日 in sorted(選んだ日):
            v = 売上.get(日)
            if not v:
                continue
            いま = 既存.get(日)
            摘要 = f"{印}（メニュー{v['menu_count']}件・商品{v['product_count']}件）"

            if いま and いま.income == v["total"]:
                飛ばした += 1
                continue
            if いま and やり方 == "skip":
                飛ばした += 1
                continue
            if いま and やり方 == "replace":
                いま.description = 摘要
                いま.income = v["total"]
                いま.save()
                直した += 1
                continue

            # 新しく入れる。空いている行があればそこへ（レシートの取り込みと同じ）
            空き = next((e for e in book.entries.all()
                        if not any([e.day, e.description, e.counter_account, e.income, e.payment])), None)
            if 空き is None:
                最後 = book.entries.last()
                空き = CashbookEntry.objects.create(book=book, row_order=(最後.row_order + 1) if 最後 else 1)
            空き.month, 空き.day = month, 日
            空き.description = 摘要
            空き.counter_account = "売上高"
            空き.income, 空き.payment = v["total"], None
            空き.save()
            入れた += 1

    return 入れた, 直した, 飛ばした
