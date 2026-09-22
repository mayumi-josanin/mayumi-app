"""経費管理（apps/keihi）— 現金出納帳（JDL 様式 777）。

いちばん間違えたくないのは**お金の計算**なので、そこを重点的に確かめる。
差引残高は行の並びで決まり、保存していないので、並べ替えれば残高も付け替わること。
書きかけの行があっても保存が止まらないこと（不備は知らせるだけ）。
"""

import pytest

from apps.keihi.models import Cashbook, CashbookEntry
from apps.keihi.services import DEFAULT_ROWS, entry_errors, running_balances, sorted_by_date, totals

pytestmark = pytest.mark.django_db

URL = "/manage/keihi/?year=2026&month=9"


def 行(**kwargs):
    """計算だけ確かめたいときの、DB に入れない身代わり。"""
    return CashbookEntry(**{"month": 9, "day": None, "income": None, "payment": None, **kwargs})


# ---- 計算 ----------------------------------------------------------------

def test_残高は前葉繰越から順に足し引きする():
    entries = [行(day=3, income=100000), 行(day=1, payment=25000), 行(day=10, payment=3500)]
    assert running_balances(50000, entries) == [150000, 125000, 121500]


def test_合計は収入と支払と最終残高():
    entries = [行(day=3, income=100000), 行(day=1, payment=25000), 行(day=10, payment=3500)]
    assert totals(50000, entries) == {
        "income_total": 100000, "payment_total": 28500, "closing_balance": 121500,
    }


def test_空の帳簿は前葉繰越がそのまま残高():
    assert totals(50000, []) == {"income_total": 0, "payment_total": 0, "closing_balance": 50000}


def test_並べ替えると残高も付け替わる():
    entries = sorted_by_date([行(day=3, income=100000), 行(day=1, payment=25000), 行(day=10, payment=3500)])
    assert [e.day for e in entries] == [1, 3, 10]
    # 1 行目が支払 25,000 に変わるので、最初の残高が変わる。最後は同じところへ着く
    assert running_balances(50000, entries) == [25000, 125000, 121500]


def test_日付の無い行は末尾へ送る():
    entries = sorted_by_date([行(day=None), 行(day=5), 行(day=2)])
    assert [e.day for e in entries] == [2, 5, None]


def test_空行は不備として知らせない():
    assert entry_errors(行()) == []


def test_収入と支払の両方に入っていたら知らせる():
    errors = entry_errors(行(day=1, income=100, payment=200))
    assert "収入金額と支払金額の両方に入っています" in errors


def test_日付が無ければ知らせる():
    assert "日付が入っていません" in entry_errors(行(description="家賃", payment=80000))


# ---- 画面 ----------------------------------------------------------------

def test_まゆみだけが入れる(client, staff):
    client.force_login(staff)
    assert client.get(URL).status_code == 403


def test_開くと紙と同じ25行がそろう(as_owner):
    assert as_owner.get(URL).status_code == 200
    book = Cashbook.objects.get(year=2026, month=9)
    assert book.entries.count() == DEFAULT_ROWS


def test_入力すると残高と合計が出る(as_owner):
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1, e2 = list(book.entries.all())[:2]

    as_owner.post(URL, {
        "opening": "50,000", "save": "1",
        f"m-{e1.id}": "9", f"d-{e1.id}": "3", f"desc-{e1.id}": "売上入金",
        f"acc-{e1.id}": "売上高", f"in-{e1.id}": "100,000", f"pay-{e1.id}": "",
        f"m-{e2.id}": "9", f"d-{e2.id}": "10", f"desc-{e2.id}": "切手",
        f"acc-{e2.id}": "通信費", f"in-{e2.id}": "", f"pay-{e2.id}": "3,500",
    })

    book.refresh_from_db()
    assert book.opening_balance == 50000
    e1.refresh_from_db()
    assert e1.income == 100000 and e1.day == 3
    assert totals(book.opening_balance, list(book.entries.all()))["closing_balance"] == 146500


def test_書きかけでも保存は止まらない(as_owner):
    """日付が無い行でも保存され、画面で知らせるだけ。手が止まらないようにしている。"""
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1 = book.entries.first()

    as_owner.post(URL, {"opening": "0", "save": "1", f"desc-{e1.id}": "家賃", f"pay-{e1.id}": "80000"})

    e1.refresh_from_db()
    assert e1.description == "家賃" and e1.payment == 80000
    assert "日付が入っていません" in entry_errors(e1)

    # 画面にも出る
    assert "確認してください" in as_owner.get(URL).content.decode()


def test_翌月は前月の最終残高を引き継ぐ(as_owner):
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1 = book.entries.first()
    as_owner.post(URL, {"opening": "50000", "save": "1",
                        f"d-{e1.id}": "3", f"in-{e1.id}": "100000"})

    as_owner.get("/manage/keihi/?year=2026&month=10")
    assert Cashbook.objects.get(year=2026, month=10).opening_balance == 150000


def test_年をまたぐときは前年12月から引き継ぐ(as_owner):
    as_owner.get("/manage/keihi/?year=2026&month=12")
    book = Cashbook.objects.get(year=2026, month=12)
    book.opening_balance = 7000
    book.save()

    as_owner.get("/manage/keihi/?year=2027&month=1")
    assert Cashbook.objects.get(year=2027, month=1).opening_balance == 7000


def test_行を足しても書きかけは消えない(as_owner):
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1 = book.entries.first()

    as_owner.post(URL, {"opening": "0", "add_row": "1", f"desc-{e1.id}": "コピー用紙"})

    e1.refresh_from_db()
    assert e1.description == "コピー用紙"
    assert book.entries.count() == DEFAULT_ROWS + 1


def test_行を消しても書きかけは消えない(as_owner):
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1, e2 = list(book.entries.all())[:2]

    as_owner.post(URL, {"opening": "0", "delete": str(e2.id), f"desc-{e1.id}": "コピー用紙"})

    e1.refresh_from_db()
    assert e1.description == "コピー用紙"
    assert not book.entries.filter(id=e2.id).exists()


def test_CSVに書き出す(as_owner):
    as_owner.get(URL)
    book = Cashbook.objects.get(year=2026, month=9)
    e1 = book.entries.first()
    as_owner.post(URL, {"opening": "0", "save": "1", f"m-{e1.id}": "9", f"d-{e1.id}": "3",
                        f"desc-{e1.id}": "売上入金", f"acc-{e1.id}": "売上高", f"in-{e1.id}": "100000"})

    r = as_owner.get("/manage/keihi/csv/?year=2026&month=9")
    本文 = r.content.decode("utf-8-sig")
    assert "日付,摘要,相手科目,収入金額,支払金額" in 本文
    assert "9/3,売上入金,売上高,100000," in 本文


def test_CSVを読み込む(as_owner):
    from django.core.files.uploadedfile import SimpleUploadedFile

    as_owner.get(URL)
    csv = "日付,摘要,相手科目,収入金額,支払金額\n9/1,コピー用紙,消耗品費,,\"25,000\"\n"
    as_owner.post("/manage/keihi/csv/import/?year=2026&month=9", {
        "year": "2026", "month": "9", "replace": "1",
        "csv_file": SimpleUploadedFile("t.csv", csv.encode("utf-8-sig"), content_type="text/csv"),
    })

    book = Cashbook.objects.get(year=2026, month=9)
    入った = book.entries.filter(description="コピー用紙").first()
    assert 入った is not None
    assert 入った.payment == 25000 and 入った.month == 9 and 入った.day == 1


def test_差引残高は出さない(as_owner):
    """院長の希望（2026-09-22）で、差引残高と前葉繰越を画面から外した。

    **数そのものは消していない**（Cashbook.opening_balance に残る）ので、戻せる。
    """
    page = as_owner.get(URL).content.decode()
    assert "差 引 残 高" not in page
    assert "前葉繰越" not in page.split("{% comment %}")[0]   # 注記の中は数えない
    assert 'id="closing-balance"' not in page
    # 収入・支払の合計は今までどおり出る
    assert 'id="income-total"' in page and 'id="payment-total"' in page


def test_繰越の欄が無くても保存できる(as_owner):
    """欄を外したので、送られてこない。**0 で上書きしない**こと。"""
    book = Cashbook.objects.create(year=2026, month=9, opening_balance=50000)
    as_owner.post(URL, {"save": "1"})
    book.refresh_from_db()
    assert book.opening_balance == 50000
