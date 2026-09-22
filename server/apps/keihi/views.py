"""現金出納帳の画面（/manage/keihi/）。

**25 行をひとつのフォームで扱う。** 連続入力の途中でページが入れ替わらないよう、
行を足す・並べ替える・行を消すもすべて同じフォームの送信にしてある。
どれを押しても、まず今の入力を保存してから、その操作をする。書きかけが消えない。

差引残高と合計は保存しない。表示のたびに services.py で計算する
（画面側でも入力中に同じ計算をして即座に出すが、正は常にこちら）。
"""

import datetime

from django.contrib import messages
from django.db import transaction
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import csv_io
from .models import Cashbook, CashbookEntry
from .services import DEFAULT_ROWS, SUGGESTED_ACCOUNTS, rows_for_display, sorted_by_date, totals


def _int_or_none(value: str) -> int | None:
    """入力欄の文字から整数を取り出す。空欄は「入れていない」として None。"""
    cleaned = (value or "").replace(",", "").strip()
    cleaned = cleaned.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if not cleaned:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def _選んだ年月(request) -> tuple[int, int]:
    today = datetime.date.today()
    try:
        year = int(request.GET.get("year") or request.POST.get("year") or today.year)
        month = int(request.GET.get("month") or request.POST.get("month") or today.month)
    except ValueError:
        return today.year, today.month
    if not 1 <= month <= 12:
        return today.year, today.month
    return year, month


def _帳簿を用意する(year: int, month: int) -> Cashbook:
    """その月の帳簿を開く。無ければ作る。

    前葉繰越は、前月の帳簿があればその最終残高を引き継ぐ。手で書き写さなくて済む。
    """
    book = Cashbook.objects.filter(year=year, month=month).first()
    if book is None:
        prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
        prev = Cashbook.objects.filter(year=prev_year, month=prev_month).first()
        opening = 0
        if prev is not None:
            opening = totals(prev.opening_balance, list(prev.entries.all()))["closing_balance"]
        book = Cashbook.objects.create(year=year, month=month, opening_balance=opening)

    _行をそろえる(book)
    return book


def _行をそろえる(book: Cashbook, minimum: int = DEFAULT_ROWS) -> None:
    """紙の様式 1 ページ分の行数までを空行で埋める。"""
    entries = list(book.entries.all())
    if len(entries) >= minimum:
        return
    order = entries[-1].row_order if entries else 0
    CashbookEntry.objects.bulk_create([
        CashbookEntry(book=book, row_order=order + i) for i in range(1, minimum - len(entries) + 1)
    ])


def _入力を保存する(request, book: Cashbook) -> None:
    """画面に出ていた行をまとめて書き戻す。"""
    # 前葉繰越の欄は画面から外した（院長の希望 2026-09-22「差引残高を削除」）。
    # **送られてこないときは、いまの値をそのまま残す。**0 で上書きすると、
    # 戻したくなったときに元の数が分からなくなる。
    if "opening" in request.POST:
        book.opening_balance = _int_or_none(request.POST.get("opening")) or 0
        book.save(update_fields=["opening_balance", "updated_at"])

    entries = list(book.entries.all())
    for entry in entries:
        entry.month = _int_or_none(request.POST.get(f"m-{entry.id}"))
        entry.day = _int_or_none(request.POST.get(f"d-{entry.id}"))
        entry.description = (request.POST.get(f"desc-{entry.id}") or "").strip()
        entry.counter_account = (request.POST.get(f"acc-{entry.id}") or "").strip()
        entry.income = _int_or_none(request.POST.get(f"in-{entry.id}"))
        entry.payment = _int_or_none(request.POST.get(f"pay-{entry.id}"))
    CashbookEntry.objects.bulk_update(
        entries, ["month", "day", "description", "counter_account", "income", "payment", "updated_at"]
    )


@owner_required
def book(request):
    year, month = _選んだ年月(request)

    if request.method == "POST":
        with transaction.atomic():
            book = _帳簿を用意する(year, month)
            _入力を保存する(request, book)

            if request.POST.get("add_row"):
                last = book.entries.last()
                CashbookEntry.objects.create(
                    book=book, row_order=(last.row_order + 1) if last else 1, month=book.month
                )
                messages.success(request, "行を足しました")
            elif request.POST.get("sort"):
                entries = sorted_by_date(list(book.entries.all()))
                for index, entry in enumerate(entries, start=1):
                    entry.row_order = index
                CashbookEntry.objects.bulk_update(entries, ["row_order"])
                messages.success(request, "日付順に並べ替えました")
            elif request.POST.get("delete"):
                book.entries.filter(id=_int_or_none(request.POST["delete"])).delete()
                messages.success(request, "行を消しました")
            else:
                messages.success(request, "保存しました")

        return redirect(f"{request.path}?year={year}&month={month}")

    book = _帳簿を用意する(year, month)
    entries = list(book.entries.all())
    前月 = datetime.date(year, month, 1) - datetime.timedelta(days=1)
    翌月 = datetime.date(year, month, 28) + datetime.timedelta(days=7)

    return render(request, "keihi/book.html", {
        "book": book,
        "rows": rows_for_display(book, entries),
        "totals": totals(book.opening_balance, entries),
        "accounts": SUGGESTED_ACCOUNTS,
        "prev": 前月,
        "next": 翌月,
        # 不備のある行だけをまとめて下に出す
        "problems": [r for r in rows_for_display(book, entries) if r["errors"]],
    })


@owner_required
def csv_export(request):
    year, month = _選んだ年月(request)
    book = _帳簿を用意する(year, month)
    content = csv_io.write_csv(list(book.entries.all()))
    filename = f"suitocho_{book.year}{book.month:02d}.csv"
    # Excel で開いたときに文字化けしないよう BOM 付きで渡す
    response = HttpResponse(content.encode("utf-8-sig"), content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f"attachment; filename={filename}"
    return response


@owner_required
@require_POST
def csv_import(request):
    year, month = _選んだ年月(request)
    uploaded = request.FILES.get("csv_file")
    if uploaded is None:
        messages.error(request, "CSV ファイルを選んでください。")
        return redirect(f"{request.path.rsplit('/', 2)[0]}/?year={year}&month={month}")

    with transaction.atomic():
        book = _帳簿を用意する(year, month)
        if request.POST.get("replace"):
            book.entries.all().delete()

        last = book.entries.last()
        order = last.row_order if last else 0
        新しい行 = []
        for i, row in enumerate(csv_io.read_csv(csv_io.decode(uploaded.read())), start=1):
            新しい行.append(CashbookEntry(book=book, row_order=order + i, **row))
        CashbookEntry.objects.bulk_create(新しい行)
        _行をそろえる(book)

    messages.success(request, f"CSV から {len(新しい行)} 行を読み込みました")
    return redirect(f"/manage/keihi/?year={year}&month={month}")
