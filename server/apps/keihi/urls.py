"""経費管理（/manage/keihi/...）。

apps/manage/urls.py から include されるので、逆引きは manage:keihi:book のように2段になる。
"""

from django.urls import path

from . import views, views_print, views_receipt

app_name = "keihi"

urlpatterns = [
    # 現金出納帳
    path("", views.book, name="book"),
    path("csv/", views.csv_export, name="csv_export"),
    path("csv/import/", views.csv_import, name="csv_import"),
    # レシート
    path("receipts/", views_receipt.receipt_list, name="receipt_list"),
    path("receipts/upload/", views_receipt.receipt_upload, name="receipt_upload"),
    path("receipts/extract/", views_receipt.receipt_extract, name="receipt_extract"),
    path("receipts/new/", views_receipt.receipt_create, name="receipt_create"),
    path("receipts/save/", views_receipt.receipt_save, name="receipt_save"),
    path("receipts/to-book/", views_receipt.receipt_to_book, name="receipt_to_book"),
    path("receipts/bulk-delete/", views_receipt.receipt_bulk_delete, name="receipt_bulk_delete"),
    path("receipts/<int:pk>/delete/", views_receipt.receipt_delete, name="receipt_delete"),
    # レシートの写真。**公開の /media/ では配らない**（管理画面の中だけ）
    path("receipts/photo/<str:名前>", views_receipt.receipt_image, name="receipt_image"),
    # 印刷（PDF はブラウザの「PDF に保存」で落とす）
    path("receipts/print/ledger/", views_print.receipt_ledger_print, name="receipt_ledger_print"),
    path("receipts/print/sheet/", views_print.receipt_sheet_print, name="receipt_sheet_print"),
]
