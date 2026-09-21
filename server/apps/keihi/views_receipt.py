"""レシートの画面（/manage/keihi/receipts/）。

写真を撮る・上げる → Claude が読む → **院長が確かめて直す** → 出納帳の行にする、という順。
読み取った時点では帳簿に入れない。AI は金額の桁や日付を読み違えることがあり、
そのまま入ると、あとから帳簿の中を探すことになるため。
"""

import datetime
import io

from django.conf import settings
from django.contrib import messages
from django.db import transaction
from django.db.models import F
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import extract, storage
from .models import Cashbook, CashbookEntry, Receipt, ReceiptItem
from .services import SUGGESTED_ACCOUNTS

# 1回の「読み取る」でまとめて読む枚数。ビジリスの分析と同じ考え方で、
# 1リクエストが長くなりすぎない数に抑える（残りはボタンをもう一度押せば続きを読む）
一度に読む枚数 = 5


def _int(value) -> int | None:
    cleaned = (value or "").replace(",", "").strip()
    cleaned = cleaned.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if not cleaned:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def _date(value) -> datetime.date | None:
    value = (value or "").strip()
    if not value:
        return None
    for 形 in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d"):
        try:
            d = datetime.datetime.strptime(value, 形).date()
            return d.replace(year=datetime.date.today().year) if 形 == "%m/%d" else d
        except ValueError:
            continue
    return None


def _一覧のURL(request) -> str:
    query = request.GET.urlencode() or request.POST.get("query", "")
    return reverse("manage:keihi:receipt_list") + (f"?{query}" if query else "")


@owner_required
def receipt_list(request):
    """レシートの一覧。既定は日付の古い順（帳簿は起きた順に見るものなので）。"""
    order = request.GET.get("order", "asc")
    受け取った = Receipt.objects.select_related("entry")
    # 日付が無いもの（読み取れなかった・これから入れる）は、どちらの順でも末尾へ送る
    if order == "desc":
        受け取った = 受け取った.order_by(F("date").desc(nulls_last=True), "-id")
    else:
        受け取った = 受け取った.order_by(F("date").asc(nulls_last=True), "id")
    受け取った = list(受け取った)

    # 分けた行（同じ写真を指す行）は、まとまりが分かるように印と合計を添える。
    # 分けた合計が元のレシートと合っているかを、目で確かめられるようにするため
    同じ写真の数: dict[str, int] = {}
    同じ写真の合計: dict[str, int] = {}
    for r in 受け取った:
        if not r.image_name:
            continue
        同じ写真の数[r.image_name] = 同じ写真の数.get(r.image_name, 0) + 1
        同じ写真の合計[r.image_name] = 同じ写真の合計.get(r.image_name, 0) + (r.amount or 0)
    行 = []
    for r in 受け取った:
        分けた数 = 同じ写真の数.get(r.image_name, 1) if r.image_name else 1
        行.append({
            "r": r,
            "分けている": 分けた数 > 1,
            "分けた数": 分けた数,
            "分けた合計": 同じ写真の合計.get(r.image_name, 0),
        })

    return render(request, "keihi/receipt_list.html", {
        "rows": 行,
        "receipts": 受け取った,
        "order": order,
        "accounts": SUGGESTED_ACCOUNTS,
        "未読み取り": sum(1 for r in 受け取った if r.status in ("pending", "failed")),
        "支払合計": sum(r.amount or 0 for r in 受け取った if r.kind == "payment"),
        "収入合計": sum(r.amount or 0 for r in 受け取った if r.kind == "income"),
        # 鍵が無いと読み取りだけができない（写真を上げて手で直すことはできる）
        "読み取りが使えない": not settings.ANTHROPIC_API_KEY,
    })


@owner_required
@require_POST
def receipt_upload(request):
    """写真をまとめて受け取る。カメラで撮ったものも、選んだファイルも同じ入口。"""
    files = request.FILES.getlist("photos")
    if not files:
        messages.error(request, "写真が選ばれていません。")
        return redirect(_一覧のURL(request))

    作った = []
    for f in files:
        try:
            名前 = storage.保存する(f)
        except Exception as e:  # 画像として開けないファイル
            messages.error(request, f"{f.name} は写真として読めませんでした（{e}）")
            continue
        作った.append(Receipt.objects.create(
            image_name=名前, original_filename=f.name, status="pending",
        ))

    messages.success(request, f"{len(作った)} 枚を受け取りました")
    # 上げた直後にそのまま読む。枚数が多いときは残りを「読み取る」で続ける
    if 作った:
        読めた, 読めなかった = _まとめて読む(作った[:一度に読む枚数])
        _読み取りの結果を伝える(request, 読めた, 読めなかった)
    return redirect(_一覧のURL(request))


def _まとめて読む(receipts) -> tuple[int, int]:
    読めた = 読めなかった = 0
    for r in receipts:
        画像 = storage.読み出す(r.image_name)
        if 画像 is None:
            r.status, r.error = "failed", "写真が見つかりませんでした"
            r.save(update_fields=["status", "error", "updated_at"])
            読めなかった += 1
            continue
        try:
            中身 = extract.読み取る(画像)
        except extract.読み取れない as e:
            r.status, r.error = "failed", str(e)
            r.save(update_fields=["status", "error", "updated_at"])
            読めなかった += 1
            continue

        r.date = _date(中身.get("date"))
        r.amount = _int(str(中身.get("amount") or ""))
        r.store_name = (中身.get("store_name") or "")[:255]
        r.counter_account = (中身.get("counter_account") or "")[:100]
        r.kind = "income" if 中身.get("kind") == "income" else "payment"
        r.memo = 中身.get("note") or ""
        r.raw = 中身.get("_raw") or ""
        r.status, r.error = "done", ""
        r.save()
        _明細を入れ直す(r, 中身.get("items"))
        読めた += 1
    return 読めた, 読めなかった


def _明細を入れ直す(r: Receipt, items) -> None:
    """読み取った品物を入れ直す。1枚の中で科目を分けるときに、これを選んでもらう。"""
    r.items.all().delete()
    if not isinstance(items, list):
        return
    作る = []
    for i, item in enumerate(items[:120], start=1):   # 長いレシート対策の上限
        if not isinstance(item, dict):
            continue
        名前 = str(item.get("name") or "").strip()[:255]
        金額 = _int(str(item.get("amount") or ""))
        if not 名前 and 金額 is None:
            continue
        作る.append(ReceiptItem(receipt=r, row_order=i, name=名前, amount=金額))
    ReceiptItem.objects.bulk_create(作る)


def _読み取りの結果を伝える(request, 読めた: int, 読めなかった: int) -> None:
    if 読めた:
        messages.success(request, f"{読めた} 枚を読み取りました。中身を確かめてから「出納帳へ」を押してください")
    if 読めなかった:
        messages.error(request, f"{読めなかった} 枚は読み取れませんでした。手で直せます")


@owner_required
@require_POST
def receipt_extract(request):
    """「読み取る」ボタン。まだ読んでいないものを、まとめて読む。"""
    待ち = list(Receipt.objects.filter(status__in=["pending", "failed"])[:一度に読む枚数])
    if not 待ち:
        messages.info(request, "読み取り待ちのレシートはありません")
        return redirect(_一覧のURL(request))
    _読み取りの結果を伝える(request, *_まとめて読む(待ち))
    return redirect(_一覧のURL(request))


@owner_required
@require_POST
def receipt_create(request):
    """写真なしで1件作る。レシートを失くした支払いや、手で足したいときに使う。"""
    r = Receipt.objects.create(status="manual", date=datetime.date.today())
    messages.success(request, "1件足しました。中身を入れて保存してください")
    return redirect(_一覧のURL(request) + f"#receipt-{r.id}")


@owner_required
def receipt_split(request, pk: int):
    """1枚のレシートを、買ったものごとに科目で分ける。

    スーパーのレシートのように、1枚の中に科目の違う買い物が混ざっていることがある。
    読み取った品物を選び、その分の科目を決めると、**選んだ品物だけが別の行に移る**。

    金額は自動で出す。**引き算で出す**のがこの画面の肝で、
      分ける行 ＝ 選んだ品物の合計
      元の行   ＝ 元の金額 − 選んだ品物の合計
    とすれば、税や値引きがどちらに入っていても、分けたあとの合計は元の金額と必ず一致する。
    品物を足し上げて作り直すと、税の分だけ帳簿が合わなくなる。

    店名・支払先・日付・種別は元のレシートから引き継ぐので、入れ直さなくてよい。
    """
    もと = get_object_or_404(Receipt, pk=pk)

    if request.method == "GET":
        return render(request, "keihi/receipt_split.html", {
            "receipt": もと,
            "items": list(もと.items.all()),
            "accounts": SUGGESTED_ACCOUNTS,
            "query": request.GET.urlencode(),
        })

    選ばれた = [_int(i) for i in request.POST.getlist("items")]
    品物 = list(もと.items.filter(id__in=[i for i in 選ばれた if i]))
    科目 = (request.POST.get("acc") or "").strip()[:100]

    # 品物が読めていないレシートは、金額を手で入れて分ける
    手入力の金額 = _int(request.POST.get("amount"))
    分ける金額 = sum(i.amount or 0 for i in 品物) if 品物 else 手入力の金額

    if not 分ける金額:
        messages.error(request, "分ける品物を選ぶか、金額を入れてください。")
        return redirect(request.path)
    if もと.amount is not None and 分ける金額 > もと.amount:
        messages.error(request, "分ける金額が、元のレシートの金額を超えています。")
        return redirect(request.path)

    with transaction.atomic():
        新しい行 = Receipt.objects.create(
            image_name=もと.image_name,        # 同じ写真を指す（消すときは最後の1行まで残す）
            original_filename=もと.original_filename,
            date=もと.date,                     # 日付・店名・種別は自動で引き継ぐ
            store_name=もと.store_name,
            kind=もと.kind,
            status=もと.status,
            amount=分ける金額,
            counter_account=科目,
            memo="、".join(i.name for i in 品物 if i.name)[:500],
        )
        # 選んだ品物は、分けた行へ移す（どちらの行に何が入っているか、あとから分かる）
        for 品 in 品物:
            品.receipt = 新しい行
            品.save(update_fields=["receipt", "updated_at"])

        if もと.amount is not None:
            もと.amount = もと.amount - 分ける金額
            残り = list(もと.items.all())
            もと.memo = "、".join(i.name for i in 残り if i.name)[:500] or もと.memo
            もと.save(update_fields=["amount", "memo", "updated_at"])

    messages.success(
        request,
        f"{分ける金額:,} 円を「{科目 or '科目なし'}」として分けました。元の行は {もと.amount:,} 円になりました"
        if もと.amount is not None else f"{分ける金額:,} 円を分けました",
    )
    return redirect(reverse("manage:keihi:receipt_list") + (f"?{request.POST.get('query', '')}" if request.POST.get("query") else ""))


@owner_required
@require_POST
def receipt_save(request):
    """一覧に出ている全部をまとめて保存する（出納帳と同じ作り）。"""
    for r in Receipt.objects.all():
        if f"date-{r.id}" not in request.POST:
            continue  # 画面に出ていなかったものは触らない
        r.date = _date(request.POST.get(f"date-{r.id}"))
        r.amount = _int(request.POST.get(f"amount-{r.id}"))
        r.store_name = (request.POST.get(f"store-{r.id}") or "").strip()[:255]
        r.counter_account = (request.POST.get(f"acc-{r.id}") or "").strip()[:100]
        r.kind = "income" if request.POST.get(f"kind-{r.id}") == "income" else "payment"
        r.memo = (request.POST.get(f"memo-{r.id}") or "").strip()
        r.save()
    messages.success(request, "保存しました")
    return redirect(_一覧のURL(request))


def _写真を片付ける(r: Receipt) -> None:
    """その写真を使っている行が他に無ければ、写真も消す。

    分けた行は同じ写真を指しているので、確かめずに消すと相方の写真が消える。
    """
    if not r.image_name:
        return
    if Receipt.objects.filter(image_name=r.image_name).exclude(pk=r.pk).exists():
        return
    storage.消す(r.image_name)


@owner_required
@require_POST
def receipt_delete(request, pk: int):
    r = get_object_or_404(Receipt, pk=pk)
    _写真を片付ける(r)
    r.delete()
    messages.success(request, "レシートを消しました（出納帳の行は残ります）")
    return redirect(_一覧のURL(request))


@owner_required
@require_POST
def receipt_bulk_delete(request):
    選ばれた = request.POST.getlist("selected")
    対象 = Receipt.objects.filter(id__in=[_int(i) for i in 選ばれた if _int(i)])
    件数 = 対象.count()
    for r in 対象:
        _写真を片付ける(r)
    対象.delete()
    messages.success(request, f"{件数} 件を消しました（出納帳の行は残ります）")
    return redirect(_一覧のURL(request))


@owner_required
@require_POST
def receipt_to_book(request):
    """選んだレシートを、その日付の月の出納帳へ行として入れる。

    **ここが「確かめてから反映」の境目。** 押されたものだけが帳簿に載る。
    """
    from .views import _帳簿を用意する

    選ばれた = [_int(i) for i in request.POST.getlist("selected")]
    対象 = Receipt.objects.filter(id__in=[i for i in 選ばれた if i], entry__isnull=True)

    入れた, 飛ばした = 0, 0
    with transaction.atomic():
        for r in 対象:
            if r.date is None or not r.amount:
                飛ばした += 1
                continue
            book = _帳簿を用意する(r.date.year, r.date.month)
            # 空いている行があればそこへ入れる。無ければ末尾に足す
            空き = next((e for e in book.entries.all()
                        if not any([e.day, e.description, e.counter_account, e.income, e.payment])), None)
            if 空き is None:
                最後 = book.entries.last()
                空き = CashbookEntry.objects.create(book=book, row_order=(最後.row_order + 1) if 最後 else 1)
            空き.month, 空き.day = r.date.month, r.date.day
            空き.description = r.store_name or r.memo or "レシート"
            空き.counter_account = r.counter_account
            if r.kind == "income":
                空き.income, 空き.payment = r.amount, None
            else:
                空き.income, 空き.payment = None, r.amount
            空き.save()
            r.entry = 空き
            r.save(update_fields=["entry", "updated_at"])
            入れた += 1

    if 入れた:
        messages.success(request, f"{入れた} 件を現金出納帳に入れました")
    if 飛ばした:
        messages.error(request, f"{飛ばした} 件は日付か金額が入っていないので入れていません")
    return redirect(_一覧のURL(request))


@owner_required
def receipt_image(request, 名前: str):
    """レシートの写真を配る。

    **公開の /media/ では配らない。** ここは管理画面の中（MANAGE_ENABLED の箱）にしか無く、
    まゆみでログインしていないと 403 になる。
    """
    中身 = storage.読み出す(名前)
    if 中身 is None:
        raise Http404
    return FileResponse(io.BytesIO(中身), content_type="image/jpeg")
