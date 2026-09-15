"""予約システムから、お客様台帳（会員）へ1行入れる窓口。

    upsertLedgerFromReservation     予約が入るたびに、予約システムのサーバーが呼ぶ

設計は docs/design/予約のお客様を台帳へ入れる.md（①②）と
docs/design/予約からLINEIDを集める.md（③の呼び水）。

## 合鍵が要る

お客様アプリではなく**予約システムのサーバー**が呼ぶ。`views.公開アクション` には
入れない（入れると誰でも会員の行を作れる）。予約システムは `.env` の合鍵を
X-Api-Key で送ってくる。

## 照らす順（設計書のとおり）

    1. LINEユーザーIDが一致（お名前も一致するとき）  → 結ぶ
    2. 氏名＋生年月日が一致                            → 結ぶ
    3. 氏名＋電話が一致                                → 結ぶ
    4. どれも無い                                      → 新しく作る（登録経路＝予約システム）

2〜3 で候補が2人以上なら**結ばない・作らない**（hold）。受付が見て決める。
機械が決めると、混ざった記録に誰も気づけない。

## ご予約なさった方と、受ける方が違うことがある

ポチコの設問に「参加される方のお名前」が実在した（代理のご予約）。
LINEユーザーIDは**ご予約なさった方**のもので、受ける方のものではない。

だから LINEユーザーID が一致しても、**お名前が違えばその ID は使わない。**
その ID の持ち主（お母様）の行に、娘さんの電話や生年月日を書き足すことになるため。
お名前が違うときは ID を無かったことにして 2〜4 へ進み、受ける方ご自身の行を
探す／作る。新しく作る行にもその ID は入れない（`lineMismatch: true` で知らせる）。

## 書くのは「空いている欄」だけ

正は台帳。予約システムの値で台帳を上書きしない。
**LINEユーザーIDが入っていて違う値なら、絶対に上書きしない。**
間違えると本当のご本人がつなげなくなり、気づくのはその方が困ったとき。
"""

import re

from django.db import transaction
from django.utils import timezone

from apps.manage import member_gate
from apps.members.models import Member

from . import admin_member, entrance

既定の経路 = "予約システム"


def _文(v):
    return "" if v is None else str(v).strip()


def _電話の数字(値):
    """比べるための電話。整えて（先頭の0を戻して）から数字だけにする。

    台帳の電話はハイフンの有無が揃っていない（受付が手で入れた行と、
    アプリから入った行が混ざっている）。そのまま比べると同じ番号が別物になる。
    """
    return re.sub(r"\D", "", admin_member._電話を整える(値))


def _生年月日(m):
    return m.birthday.isoformat() if m.birthday else ""


def _答え(result, m=None, matched_by=None, filled=None, candidates=0, **追加):
    答 = {
        "status": "ok",
        "result": result,
        "memberId": m.member_id if m else None,
        "matchedBy": matched_by,
        "filled": list(filled or []),
        "candidates": candidates,
    }
    答.update(追加)
    return 答


def 予約から入れる(d):
    """`upsertLedgerFromReservation`。受け取る項目:

        name（必須） kana phone birthday(YYYY-MM-DD) address lineUserId email note
        reservationId source（既定「予約システム」）

    `email` は台帳（Member）に置き場が無いので、受け取るが書かない。
    `note` は新しく作る行のメモにだけ足す（結んだ行のメモは触らない。正は台帳）。
    """
    # 9/19 まで会員の正はスプレッドシート。**それまでにサーバーへ書くと正が2つになる。**
    # 予約側は retryLater を見て、切り替えのあとに送り直す（booking.Customer に全部残っている）。
    if not member_gate.会員はサーバーが正():
        return {"status": "error",
                "message": "会員の表はまだスプレッドシートが正です。9/19 の切り替え後に送り直してください。",
                "retryLater": True}

    生の名 = _文(d.get("name"))
    名 = entrance._名前で照らす(生の名)
    if not 名:
        return {"status": "error", "message": "お名前が空です。予約システムの記録を確かめてください。"}

    かな = _文(d.get("kana"))
    電話 = _文(d.get("phone"))
    生年月日 = entrance._日付の文字(d.get("birthday"))
    住所 = _文(d.get("address"))
    LINE = _文(d.get("lineUserId"))
    予約ID = _文(d.get("reservationId"))
    経路 = _文(d.get("source")) or 既定の経路
    覚え書き = _文(d.get("note"))

    with transaction.atomic():
        # 同時に届いても同じ行を2つ作らないよう、新規登録と同じく全行を押さえる。
        全員 = list(Member.objects.select_for_update().exclude(deleted=True))

        m = None
        照合 = None
        LINEは使えない = False  # ご本人の ID ではない（別の名前の行に付いている）と分かったとき

        # (1) LINEユーザーID
        if LINE:
            持ち主 = [x for x in 全員 if x.line_user_id == LINE]  # unique なので最大1人
            if 持ち主 and entrance._名前で照らす(持ち主[0].name) == 名:
                m, 照合 = 持ち主[0], "line"
            elif 持ち主:
                LINEは使えない = True

        # (2) 氏名＋生年月日
        if m is None and 生年月日:
            候補 = [x for x in 全員 if entrance._名前で照らす(x.name) == 名 and _生年月日(x) == 生年月日]
            if len(候補) > 1:
                return _答え("hold", candidates=len(候補), reason="nameBirthday")
            if 候補:
                m, 照合 = 候補[0], "nameBirthday"

        # (3) 氏名＋電話
        if m is None and 電話:
            数字 = _電話の数字(電話)
            候補 = [x for x in 全員
                    if 数字 and entrance._名前で照らす(x.name) == 名 and _電話の数字(x.phone) == 数字]
            if len(候補) > 1:
                return _答え("hold", candidates=len(候補), reason="namePhone")
            if 候補:
                m, 照合 = 候補[0], "namePhone"

        # LINE を書いてよいか。unique なので、退会者を含めて他の行に同じ ID があれば書けない。
        # (1) で当たらなかったのに他の行が持っている＝退会した行か、別の名前の行。
        LINEを書く = ""
        LINE不一致 = False
        if LINE and not LINEは使えない:
            他 = Member.objects.filter(line_user_id=LINE)
            if m is not None:
                他 = 他.exclude(pk=m.pk)
            if 他.exists():
                # 別の行がその ID を持っている。**結ばない・作らない。**受付が見る。
                return _答え("hold", candidates=1, reason="lineTaken")
            LINEを書く = LINE
        if LINEは使えない:
            LINE不一致 = True

        埋めた = []

        if m is not None:
            # 結んだ行の**空いている欄だけ**埋める。
            if かな and not (m.kana or "").strip():
                m.kana = entrance._保存するかな(かな)
                埋めた.append("kana")
            if 電話 and not (m.phone or "").strip():
                m.phone = admin_member._電話を整える(電話)
                埋めた.append("phone")
            if 生年月日 and not m.birthday:
                m.birthday = entrance.parse_date(生年月日)
                埋めた.append("birthday")
            if 住所 and not (m.address or "").strip():
                m.address = 住所
                埋めた.append("address")
            if LINEを書く:
                if not (m.line_user_id or "").strip():
                    m.line_user_id = LINEを書く
                    埋めた.append("lineUserId")
                elif m.line_user_id != LINEを書く:
                    # **入っていて違う値なら上書きしない。**
                    LINE不一致 = True
            if 埋めた:
                m.save()
            return _答え("linked", m, 照合, 埋めた, lineMismatch=LINE不一致)

        # (4) 新しく作る
        会員ID = entrance.会員IDを採番する()
        if not 会員ID:
            # **黙って番号を作らない。**予約側は error を記録し、次のご予約で送り直す。
            return {"status": "error",
                    "message": "会員番号を採番できませんでした。受付にお申し出ください。"}

        今 = timezone.now()
        メモ = f"予約システムから取り込み（{timezone.localdate(今).isoformat()}）"
        if 覚え書き:
            メモ += "\n" + 覚え書き
        m = Member.objects.create(
            member_id=会員ID,
            created_at=今,
            name=entrance._保存する名前(生の名),
            kana=entrance._保存するかな(かな) if かな else "",
            phone=admin_member._電話を整える(電話) if 電話 else "",
            birthday=entrance.parse_date(生年月日) if 生年月日 else None,
            address=住所,
            line_user_id=LINEを書く or None,
            memo=メモ,
            # パスコードは空。その方はまだアプリにログインできない（登録のときに引き受ける）。
            passcode_hash="",
            device_sessions=[],
            registration_source=経路,
            registration_source_detail=予約ID,
            registration_source_updated_at=今,
        )
        埋めた = [k for k, v in (("kana", かな), ("phone", 電話), ("birthday", 生年月日),
                                ("address", 住所), ("lineUserId", LINEを書く)) if v]
        return _答え("created", m, None, 埋めた, lineMismatch=LINE不一致)
