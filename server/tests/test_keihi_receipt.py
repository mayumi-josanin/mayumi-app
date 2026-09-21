"""経費管理のレシート（apps/keihi）。

**いちばん確かめたいのは「読んだだけでは帳簿に入らない」こと。**
AI は金額の桁や日付を読み違えるので、院長が押したものだけが出納帳の行になる。
Claude への通信は差し替えて、お金の流れだけを見る。
"""

import datetime
import io

import pytest
from django.core.files.uploadedfile import SimpleUploadedFile

from apps.keihi import extract, storage
from apps.keihi.models import Cashbook, Receipt

pytestmark = pytest.mark.django_db

URL = "/manage/keihi/receipts/"


@pytest.fixture(autouse=True)
def _置き場(settings, tmp_path):
    settings.MEDIA_ROOT = tmp_path / "media"


def 写真(name="receipt.jpg"):
    """PIL が開ける、小さな本物の JPEG。"""
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (40, 60), "white").save(buf, "JPEG")
    return SimpleUploadedFile(name, buf.getvalue(), content_type="image/jpeg")


@pytest.fixture
def 読める(monkeypatch):
    """Claude が読み取れたことにする。"""
    def fake(画像):
        return {
            "date": "2026-09-03", "amount": 1280, "store_name": "まるみ文具店",
            "counter_account": "消耗品費", "kind": "payment", "note": "コピー用紙",
            "_raw": "{}",
        }
    monkeypatch.setattr(extract, "読み取る", fake)


@pytest.fixture
def 読めない(monkeypatch):
    def fake(画像):
        raise extract.読み取れない("ぼやけていて読めません")
    monkeypatch.setattr(extract, "読み取る", fake)


# ---- 上げる・読む ---------------------------------------------------------

def test_まゆみだけが入れる(client, staff):
    client.force_login(staff)
    assert client.get(URL).status_code == 403


def test_複数枚まとめて上げられる(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真("a.jpg"), 写真("b.jpg"), 写真("c.jpg")]})
    assert Receipt.objects.count() == 3


def test_上げると読み取って項目が入る(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    assert r.status == "done"
    assert r.date == datetime.date(2026, 9, 3)
    assert r.amount == 1280
    assert r.store_name == "まるみ文具店"
    assert r.counter_account == "消耗品費"


def test_読めなくても消えずに残り手で直せる(as_owner, 読めない):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    assert r.status == "failed"
    assert "ぼやけて" in r.error

    as_owner.post(f"{URL}save/", {
        f"date-{r.id}": "2026-09-05", f"amount-{r.id}": "3,500",
        f"store-{r.id}": "郵便局", f"acc-{r.id}": "通信費", f"kind-{r.id}": "payment",
        f"memo-{r.id}": "切手",
    })
    r.refresh_from_db()
    assert r.amount == 3500 and r.store_name == "郵便局"


def test_写真として読めないファイルは断る(as_owner, 読める):
    悪い = SimpleUploadedFile("x.jpg", b"not an image", content_type="image/jpeg")
    as_owner.post(f"{URL}upload/", {"photos": [悪い]})
    assert Receipt.objects.count() == 0


# ---- 確かめてから反映 -----------------------------------------------------

def test_読んだだけでは出納帳に入らない(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    assert Receipt.objects.get().entry_id is None
    assert Cashbook.objects.count() == 0


def test_出納帳へ押したものだけが行になる(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真("a.jpg"), 写真("b.jpg")]})
    a, b = Receipt.objects.order_by("id")

    as_owner.post(f"{URL}to-book/", {"selected": [str(a.id)]})

    a.refresh_from_db(); b.refresh_from_db()
    assert a.entry_id is not None
    assert b.entry_id is None

    行 = a.entry
    assert (行.month, 行.day) == (9, 3)
    assert 行.payment == 1280 and 行.income is None
    assert 行.counter_account == "消耗品費"
    assert 行.description == "まるみ文具店"
    assert 行.book.year == 2026 and 行.book.month == 9


def test_収入のレシートは収入欄に入る(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    as_owner.post(f"{URL}save/", {
        f"date-{r.id}": "2026-09-03", f"amount-{r.id}": "1280",
        f"store-{r.id}": "売上", f"acc-{r.id}": "売上高", f"kind-{r.id}": "income", f"memo-{r.id}": "",
    })
    as_owner.post(f"{URL}to-book/", {"selected": [str(r.id)]})

    r.refresh_from_db()
    assert r.entry.income == 1280 and r.entry.payment is None


def test_日付か金額が無いものは入れない(as_owner, 読めない):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    as_owner.post(f"{URL}to-book/", {"selected": [str(r.id)]})
    r.refresh_from_db()
    assert r.entry_id is None


def test_二度押しても二重に入らない(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    as_owner.post(f"{URL}to-book/", {"selected": [str(r.id)]})
    as_owner.post(f"{URL}to-book/", {"selected": [str(r.id)]})

    book = Cashbook.objects.get()
    書かれた行 = [e for e in book.entries.all() if e.payment]
    assert len(書かれた行) == 1


# ---- 作る・消す・並べる ---------------------------------------------------

def test_写真なしで1件作れる(as_owner):
    as_owner.post(f"{URL}new/")
    r = Receipt.objects.get()
    assert r.status == "manual" and r.image_name == ""


def test_消しても出納帳の行は残る(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    as_owner.post(f"{URL}to-book/", {"selected": [str(r.id)]})
    r.refresh_from_db()
    行id = r.entry_id

    as_owner.post(f"{URL}{r.id}/delete/")

    assert Receipt.objects.count() == 0
    from apps.keihi.models import CashbookEntry
    assert CashbookEntry.objects.get(id=行id).payment == 1280


def test_消すと写真もサーバーから消える(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    r = Receipt.objects.get()
    assert storage.読み出す(r.image_name) is not None

    as_owner.post(f"{URL}{r.id}/delete/")
    assert storage.読み出す(r.image_name) is None


def test_まとめて消せる(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真("a.jpg"), 写真("b.jpg")]})
    ids = [str(r.id) for r in Receipt.objects.all()]
    as_owner.post(f"{URL}bulk-delete/", {"selected": ids})
    assert Receipt.objects.count() == 0


def test_古い順と新しい順(as_owner):
    Receipt.objects.create(date=datetime.date(2026, 9, 10), store_name="あと", status="manual")
    Receipt.objects.create(date=datetime.date(2026, 9, 1), store_name="さき", status="manual")
    Receipt.objects.create(store_name="日付なし", status="manual")

    古い順 = as_owner.get(URL).context["receipts"]
    assert [r.store_name for r in 古い順] == ["さき", "あと", "日付なし"]

    新しい順 = as_owner.get(f"{URL}?order=desc").context["receipts"]
    assert [r.store_name for r in 新しい順] == ["あと", "さき", "日付なし"]


# ---- 写真の見せ方（外から見えないこと） -----------------------------------

def test_写真はログインしていないと見られない(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})
    名前 = Receipt.objects.get().image_name

    # as_owner は conftest の client そのものなので、別の入れ物で「よその人」を作る
    from django.test import Client
    よその人 = Client()
    assert よその人.get(f"{URL}photo/{名前}").status_code in (302, 403)

    assert as_owner.get(f"{URL}photo/{名前}").status_code == 200


def test_変な名前で外のファイルを読ませない(as_owner):
    assert as_owner.get(f"{URL}photo/..%2F..%2Fsettings.py").status_code == 404


# ---- 印刷 -----------------------------------------------------------------

def test_台帳と台紙が出る(as_owner, 読める):
    as_owner.post(f"{URL}upload/", {"photos": [写真()]})

    台帳 = as_owner.get(f"{URL}print/ledger/")
    assert 台帳.status_code == 200
    assert "まるみ文具店" in 台帳.content.decode()

    台紙 = as_owner.get(f"{URL}print/sheet/")
    assert 台紙.status_code == 200
    assert "レ シ ー ト 台 紙" in 台紙.content.decode()
