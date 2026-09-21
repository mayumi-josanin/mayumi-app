"""印刷の画面（PDF はブラウザの「PDF に保存」で落とす）。

外の PDF 部品は入れていない。日本語の字と、貼り込むレシートの写真が、
いちばん確実に出るのがブラウザの印刷だから。押すと印刷の窓が開くようにしてある。

出すのは3つ：
  現金出納帳（紙の様式）… views.book の画面をそのまま印刷（ここでは扱わない）
  レシートの台帳（一覧表）… 日付・店名・金額・相手科目を並べた表
  レシートの台紙（写真貼り）… 写真を1枚ずつ貼り、日付と金額を添えたもの
"""

from django.db.models import F
from django.shortcuts import render

from apps.manage.permissions import owner_required

from .models import Receipt


def _並べる(request):
    受け取った = Receipt.objects.select_related("entry")
    if request.GET.get("order") == "desc":
        return list(受け取った.order_by(F("date").desc(nulls_last=True), "-id"))
    return list(受け取った.order_by(F("date").asc(nulls_last=True), "id"))


@owner_required
def receipt_ledger_print(request):
    """レシートの台帳（一覧表）。税理士さんへ渡す形。"""
    受け取った = _並べる(request)
    return render(request, "keihi/receipt_ledger_print.html", {
        "receipts": 受け取った,
        "支払合計": sum(r.amount or 0 for r in 受け取った if r.kind == "payment"),
        "収入合計": sum(r.amount or 0 for r in 受け取った if r.kind == "income"),
    })


@owner_required
def receipt_sheet_print(request):
    """レシートの台紙。写真を貼って日付と金額を添えたもの。紙で保存する用。"""
    受け取った = [r for r in _並べる(request) if r.image_name]
    return render(request, "keihi/receipt_sheet_print.html", {"receipts": 受け取った})
