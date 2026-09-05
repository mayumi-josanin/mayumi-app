"""カテゴリマスタを、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**（[窓口はGASのまま](docs/design/窓口はGASのまま.md)）

## なぜお知らせの次にこれなのか

お知らせをサーバーへ移した時点で、**カテゴリの食い違いが生まれていた。**
GAS の `getNews` はカテゴリの一覧も一緒に返すが、お知らせがサーバーから
返るようになったので、**そこに載るカテゴリもサーバーの写し**になる。
院長がカテゴリを足しても、シートに入るだけでお客様の画面には出ない。

## 名前が鍵

カテゴリは行番号ではなく**名前**で指される（GAS の handleUpdateCategory は
oldName で探す）。お知らせ・商品・メニューが名前の文字列で参照しているため。
**IDに直したくなるが、直すと参照元をすべて書き換えることになる。**
移行のあいだは名前のまま置く。

## 名前を変えても、記事のカテゴリは変えない

GAS もそうしている。名前を変えると、その名前を使っていた記事は
**古い名前のまま取り残される。**今の振る舞いをそのまま写す。
直すなら切り替えが済んでから。
"""

from django.db import transaction

from apps.content.models import Category

種別 = ["お知らせ", "ブログ", "メニュー", "通知"]


def _文(v):
    return "" if v is None else str(v)


def _種別(v):
    """GAS と同じ。**当てはまらなければ「ブログ」。**"""
    t = _文(v).strip()
    return t if t in 種別 else "ブログ"


def 一覧():
    """GAS の getCategories と同じ。**同じ名前は最初の1つだけ。**"""
    出, 見た = [], set()
    for c in Category.objects.all().order_by("sheet_row"):
        名 = (c.name or "").strip()
        if not 名 or 名 in 見た:
            continue
        見た.add(名)
        出.append({"name": 名, "type": _種別(c.kind)})
    return {"status": "ok", "categories": 出}


def 足す(d):
    """GAS の handleAddCategory。**同じ名前があれば断る。**"""
    名 = _文(d.get("name")).strip()
    if not 名:
        return {"status": "error", "message": "カテゴリ名を入力してください"}

    with transaction.atomic():
        if Category.objects.filter(name=名).exists():
            return {"status": "error", "message": "同じカテゴリ名が既に登録されています"}
        最大 = Category.objects.select_for_update().order_by("-sheet_row").values_list(
            "sheet_row", flat=True).first() or 1
        Category.objects.create(sheet_row=最大 + 1, name=名, kind=_種別(d.get("categoryType")))
    return {"status": "ok", "message": "カテゴリを追加しました"}


def 書き換える(d):
    """GAS の handleUpdateCategory。**oldName で探す。**"""
    旧 = _文(d.get("oldName")).strip()
    新 = _文(d.get("newName")).strip()
    if not 旧:
        return {"status": "error", "message": "更新対象のカテゴリが見つかりません"}
    if not 新:
        return {"status": "error", "message": "カテゴリ名を入力してください"}

    with transaction.atomic():
        c = Category.objects.select_for_update().filter(name=旧).first()
        if not c:
            return {"status": "error", "message": "更新対象のカテゴリが見つかりません"}
        if 新 != 旧 and Category.objects.filter(name=新).exists():
            return {"status": "error", "message": "同じカテゴリ名が既に登録されています"}
        c.name = 新
        c.kind = _種別(d.get("categoryType"))
        c.save(update_fields=["name", "kind", "changed_at"])
    return {"status": "ok", "message": "カテゴリを更新しました"}


def 消す(d):
    """GAS の handleDeleteCategory。**同じ名前が複数あれば全部消す。**

    GAS は名前の一致する行をすべて集めて消している。ここも同じにする。
    行番号は繰り上げない。**カテゴリは名前で指されるので、番号は誰も見ていない。**
    （FAQ は行番号で指されるので繰り上げた。表ごとに違う。）
    """
    名 = _文(d.get("name")).strip()
    if not 名:
        return {"status": "error", "message": "削除対象のカテゴリが見つかりません"}

    with transaction.atomic():
        件 = Category.objects.select_for_update().filter(name=名)
        数 = 件.count()
        if not 数:
            return {"status": "error", "message": "削除対象のカテゴリが見つかりません"}
        件.delete()
    return {"status": "ok", "message": "カテゴリを削除しました", "deleted": 数}
