"""使い方FAQを、管理アプリから読み書きするための窓口。

**GAS が転送してくる先。**（[窓口はGASのまま](docs/design/窓口はGASのまま.md)）

## 行番号のずれ方まで、GAS に合わせる

GAS の `handleDeleteSupportFaq` は `sheet.deleteRow()` で**本当に行を消す。**
消すと、それより下の行番号が1つずつ繰り上がる。

    行5 を消す → もとの行6が行5になる

ここでも同じにする。**残して印を付ける形にはしない。**
お知らせは論理削除だが、FAQ は違う。表ごとに今の振る舞いが違うので、
**「サーバー側の作法」でそろえず、その表の今の振る舞いに合わせる。**
切り替えで見た目や番号の付き方が変わると、原因の切り分けができなくなる。

そのぶん、**消したものは戻らない。**日次のデータベースの控えが頼りになる。
GAS のいまの振る舞いも同じなので、危うさが増えたわけではない。
"""

from django.db import transaction
from django.db.models import F
from django.utils import timezone

from apps.content.models import SupportFaq


def _文(v):
    return "" if v is None else str(v)


def _時刻の字(d):
    """GAS が返している形（`2026/4/4 0:32`）にそろえる。

    FAQ は `getDisplayValues()` で読んでいるので、**シートの表示そのまま**が
    返る。ISO の形にすると管理画面の見え方が変わる。
    移行では見た目を変えない。直すなら切り替えが済んでから。

    月・日・時は**先頭の0を付けない。**分は付ける（シートの表示がそう）。
    """
    if not d:
        return ""
    t = timezone.localtime(d)
    return f"{t.year}/{t.month}/{t.day} {t.hour}:{t.minute:02d}"


def _一件(f):
    """GAS の getSupportFaqEntries_ が返す1件と同じ形。"""
    return {
        "rowIdx": f.sheet_row,
        "status": "公開" if f.published else "非公開",
        "category": f.category or "",
        "question": f.question or "",
        "keywords": f.keywords or "",
        "answer": f.answer or "",
        "priority": f.priority or 0,
        "updatedAt": _時刻の字(f.updated_at),
    }


def 一覧():
    """GAS の getAdminSupportFaq。**非公開のものも返す。**

    GAS は「質問と回答が両方ある」ものだけを出し、
    **優先度の高い順、同点はシートの行の順**に並べる。そこまで同じにする。
    """
    件 = [
        _一件(f)
        for f in SupportFaq.objects.exclude(question="").exclude(answer="")
    ]
    件.sort(key=lambda x: (-x["priority"], x["rowIdx"]))
    return {"status": "ok", "faqs": 件}


def 保存(d):
    """GAS の handleSaveSupportFaq。**rowIdx があれば更新、無ければ追加。**"""
    質問 = _文(d.get("question")).strip()
    回答 = _文(d.get("answer")).strip()
    if not 質問:
        return {"status": "error", "message": "質問を入力してください"}
    if not 回答:
        return {"status": "error", "message": "回答を入力してください"}

    値 = {
        "published": _文(d.get("status")).strip() != "非公開",
        "category": _文(d.get("category")).strip(),
        "question": 質問,
        "keywords": _文(d.get("keywords")).strip(),
        "answer": 回答,
        "priority": _数(d.get("priority")),
        "updated_at": timezone.now(),
    }

    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        行 = 0

    with transaction.atomic():
        if 行 > 1:
            f = SupportFaq.objects.select_for_update().filter(sheet_row=行).first()
            if not f:
                return {"status": "error", "message": "FAQが見つかりませんでした"}
            for k, v in 値.items():
                setattr(f, k, v)
            f.save()
            return {"status": "ok", "message": "FAQを更新しました"}

        # 追加。**いまの最大＋1。**シートの appendRow と同じ位置になる。
        最大 = SupportFaq.objects.select_for_update().order_by(
            "-sheet_row").values_list("sheet_row", flat=True).first() or 1
        SupportFaq.objects.create(sheet_row=最大 + 1, **値)
    return {"status": "ok", "message": "FAQを追加しました"}


def _数(v):
    try:
        return int(float(v or 0))
    except (TypeError, ValueError):
        return 0


def 消す(d):
    """GAS の handleDeleteSupportFaq。**本当に消し、下の行番号を繰り上げる。**

    シートの deleteRow と同じ振る舞い。繰り上げないと、管理アプリが
    次に取り直したときの番号と、こちらの番号がずれていく。
    """
    try:
        行 = int(d.get("rowIdx") or 0)
    except (TypeError, ValueError):
        行 = 0
    if 行 <= 1:
        return {"status": "error", "message": "削除対象が見つかりません"}

    with transaction.atomic():
        f = SupportFaq.objects.select_for_update().filter(sheet_row=行).first()
        if not f:
            return {"status": "error", "message": "削除対象が見つかりません"}
        f.delete()
        # **下から順に繰り上げる。**一括で -1 すると、sheet_row が unique なので
        # 途中で衝突する可能性がある。番号の小さいほうから詰めれば衝突しない。
        for other in SupportFaq.objects.filter(sheet_row__gt=行).order_by("sheet_row"):
            SupportFaq.objects.filter(pk=other.pk).update(sheet_row=F("sheet_row") - 1)
    return {"status": "ok", "message": "FAQを削除しました"}
