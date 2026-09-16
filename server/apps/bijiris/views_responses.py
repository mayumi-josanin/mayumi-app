"""ビジリス管理「回答管理」。**中身は旧ビジリス管理アプリ（#page-responses、app.js renderResponses 5713行〜）と同じ。**

一覧（response_list）: 絞り込み（アンケート・対応状況・検索・会員番号・フリガナ・管理メモ・お悩みカテゴリ・回数券・何枚目・
何回目・写真・写真枚数・開始日・終了日、未対応のみ・CSV出力・PDF出力）→ アンケートタイトル一覧 → 回答者の一覧（最新順）。
詳細（response_detail）: 前回比較・写真比較・回答の全項目（直せる）・アップロード写真・対応状況と管理メモの保存・
ゴミ箱へ移動／完全削除・印刷。

保存の規則は Code.gs の updateResponse_（4677行〜。normalizeAdminAnswers_・normalizeStatus_）と同じ。
**ゴミ箱は対応状況 status="trash"。**旧アプリの「完全削除」は adminDelete → trashResponse_（実はゴミ箱へ入れるだけ）で、
本当に消すのは purgeResponse_（保管日数を過ぎたものの掃除）。ここでは画面の文言どおり、ゴミ箱の回答だけ本当に消す。

書くのは切り替えの印（gate）が立っているときだけ。立つまでは断って一覧へ戻す。
"""

import csv
import io
from datetime import datetime, time
from datetime import timezone as dt_timezone
from pathlib import Path
from urllib.parse import urlencode

from django.conf import settings
from django.contrib import messages
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import answers, gate
from .models import CustomerProfile, Response, Survey

絞り込みの鍵 = ["survey", "status", "q", "member", "kana", "memo", "concern", "plan", "sheet", "round",
           "photo", "photos", "from", "to", "group"]
# 送信日時が空の回答（取り込みで読めなかったもの）を並べるときの仮の値。一番古い扱い
昔 = datetime(2000, 1, 1, tzinfo=dt_timezone.utc)


def _絞り込みの値(request) -> dict:
    return {k: (request.GET.get(k) or "").strip() for k in 絞り込みの鍵}


def _絞り込みの文字列(値: dict, **上書き) -> str:
    d = {k: v for k, v in {**値, **上書き}.items() if v}
    return urlencode(d)


def _写真枚数の区分(n: int) -> str:
    if n >= 5:
        return "5plus"
    if n >= 3:
        return "3-4"
    if n >= 1:
        return "1-2"
    return "0"


def _日の境(文字, 終わり=False):
    """開始日・終了日（YYYY-MM-DD、日本時間）。終了日は 23:59:59。"""
    try:
        d = datetime.strptime(文字, "%Y-%m-%d").date()
    except ValueError:
        return None
    t = time(23, 59, 59) if 終わり else time(0, 0, 0)
    return timezone.make_aware(datetime.combine(d, t), timezone.get_current_timezone())


def _全部():
    return list(Response.objects.all().prefetch_related("photos").select_related("survey"))


def 絞り込む(responses, 値: dict, 控え: answers.顧客の控え) -> list:
    """getFilteredResponses と同じ。対応状況が空なら**ゴミ箱を除く**。新しい順。"""
    出 = []
    開始 = _日の境(値["from"]) if 値["from"] else None
    終了 = _日の境(値["to"], True) if 値["to"] else None
    語 = 値["q"].lower()
    会員番号 = 値["member"].upper()
    かな = answers.ゆるい文字(値["kana"])
    for r in responses:
        状況 = answers.対応状況を整える(r.status)
        if 値["survey"] and r.survey_key != 値["survey"]:
            continue
        if 値["status"]:
            if 状況 != 値["status"]:
                continue
        elif 状況 == "trash":
            continue
        if 開始 and (not r.submitted_at or r.submitted_at < 開始):
            continue
        if 終了 and (not r.submitted_at or r.submitted_at > 終了):
            continue
        if 語 and 語 not in answers.検索用の文(r, 控え):
            continue
        if 会員番号 and 会員番号 not in 控え.会員番号(r):
            continue
        if かな and かな not in answers.ゆるい文字(控え.フリガナ(r)):
            continue
        メモ = (r.admin_memo or "").strip()
        if 値["memo"] == "with" and not メモ:
            continue
        if 値["memo"] == "without" and メモ:
            continue
        if 値["concern"] and 値["concern"] not in answers.お悩みの分類ID(r):
            continue
        券 = dict(answers.回数券情報(r))
        if 値["plan"] and 券.get("回数券") != 値["plan"]:
            continue
        if 値["sheet"] and 券.get("何枚目") != 値["sheet"]:
            continue
        if 値["round"] and 券.get("何回目") != 値["round"]:
            continue
        枚 = len(answers.写真の一覧(r))
        if 値["photo"] == "with" and not 枚:
            continue
        if 値["photo"] == "without" and 枚:
            continue
        if 値["photos"] and _写真枚数の区分(枚) != 値["photos"]:
            continue
        出.append(r)
    出.sort(key=lambda r: r.submitted_at or 昔, reverse=True)
    return 出


def アンケートごとに(responses) -> list[dict]:
    """groupResponsesBySurvey: 鍵は回答の surveyId（無ければアンケート名）。最新の回答が新しい順。"""
    群 = {}
    for r in responses:
        鍵 = r.survey_key or r.survey_title or r.response_id
        g = 群.setdefault(鍵, {"key": 鍵, "title": r.survey_title or "アンケート", "latest": r.submitted_at, "count": 0,
                              "unread": 0, "responses": [], "names": set()})
        g["responses"].append(r)
        g["count"] += 1
        g["names"].add(r.customer_name)
        if answers.対応状況を整える(r.status) == "new":
            g["unread"] += 1
        if r.submitted_at and (not g["latest"] or r.submitted_at > g["latest"]):
            g["latest"] = r.submitted_at
    出 = sorted(群.values(), key=lambda g: g["latest"] or 昔, reverse=True)
    for g in 出:
        g["respondents"] = len(g["names"])
        g["latest_label"] = answers.日時の文字(g["latest"])
    return 出


def _画面の共通(title="回答管理"):
    return {"title": title, "server_is_source": gate.ビジリスはサーバーが正(), "refuse_text": gate.断る文()}


def _状況の表示(r):
    s = answers.対応状況を整える(r.status)
    return {"status": s, "status_label": answers.対応状況の名[s], "badge": answers.対応状況の印[s]}


@owner_required
def response_list(request):
    値 = _絞り込みの値(request)
    控え = answers.顧客の控え(CustomerProfile.objects.all())
    絞った = 絞り込む(_全部(), 値, 控え)
    群 = アンケートごとに(絞った)
    選んだ群 = next((g for g in 群 if g["key"] == 値["group"]), None) if 値["group"] else None
    if 値["group"] and not 選んだ群:
        値["group"] = ""
    一覧 = []
    if 選んだ群:
        for r in sorted(選んだ群["responses"], key=lambda r: r.submitted_at or 昔, reverse=True):
            一覧.append({"r": r, **_状況の表示(r), "delete_label": "完全削除" if r.trashed else "回答削除"})
    return render(request, "bijiris/response_list.html", {
        **_画面の共通(),
        "f": 値,
        "surveys": Survey.objects.all(),
        "concern_categories": answers.お悩みの分類,
        "plans": answers.回数券の種類, "sheets": answers.回数券の枚目, "rounds": answers.回数券の回目,
        "groups": 群,
        "selected_group": 選んだ群,
        "items": 一覧,
        "query": _絞り込みの文字列(値),
        "query_without_group": _絞り込みの文字列(値, group=""),
        "query_unread": _絞り込みの文字列(値, status="new", group=""),
    })


# ---------------------------------------------------------------------------
# 詳細
# ---------------------------------------------------------------------------

def _同じ人の同じアンケート(r, 全部):
    return [x for x in 全部 if x.customer_name == r.customer_name and x.survey_key == r.survey_key and not x.trashed]


def 前回の回答(r, 全部):
    候補 = [x for x in _同じ人の同じアンケート(r, 全部)
            if x.pk != r.pk and x.submitted_at and r.submitted_at and x.submitted_at < r.submitted_at]
    return max(候補, key=lambda x: x.submitted_at) if 候補 else None


def _値の文(answer, 写真の表) -> str:
    枚 = len(写真の表.get(answer.get("questionId"), []))
    if 枚:
        return f"{枚}枚"
    return answers.文字(answer.get("value")) or "未回答"


def 前回比較(r, survey, 全部) -> dict:
    """renderComparisonSection: 設問ごとに前回と違うところだけ。"""
    前 = 前回の回答(r, 全部)
    if not survey or not 前:
        return {"previous": None, "rows": []}
    今の表 = {a["questionId"]: a for a in answers.表示用の回答(r, survey)}
    前の表 = {a["questionId"]: a for a in answers.表示用の回答(前, survey)}
    今の写真 = answers.写真を設問ごとに(r)
    前の写真 = answers.写真を設問ごとに(前)
    行 = []
    for q in survey.questions or []:
        if not isinstance(q, dict):
            continue
        a = 今の表.get(q.get("id")) or {"questionId": q.get("id"), "value": ""}
        b = 前の表.get(q.get("id")) or {"questionId": q.get("id"), "value": ""}
        今値, 前値 = _値の文(a, 今の写真), _値の文(b, 前の写真)
        if 今値 != 前値:
            行.append({"label": q.get("label") or "", "required": q.get("required") is not False, "last": 前値, "current": 今値})
    return {"previous": 前, "previous_label": answers.日時の文字(前.submitted_at), "rows": 行}


def _写真の群(r, survey, 控え) -> list[dict]:
    """写真のある回答項目ごとに（設問の見出し + 写真）。"""
    写真 = answers.写真を設問ごとに(r)
    出 = []
    for a in answers.表示用の回答(r, survey):
        if 写真.get(a["questionId"]):
            出.append({"label": a["label"], "photos": 写真[a["questionId"]]})
    # 設問に結びつかない写真（取り込みで questionId が空だったもの）も落とさない
    余り = [p for qid, ps in 写真.items() if qid not in {a["questionId"] for a in answers.表示用の回答(r, survey)} for p in ps]
    if 余り:
        出.append({"label": "写真", "photos": 余り})
    return 出


def 写真比較(r, survey, 全部, 控え) -> dict:
    """getPhotoComparisonSlots + renderPhotoComparisonSection: 初回・前回・今回。"""
    if not survey:
        return {"slots": []}
    写真あり = sorted([x for x in _同じ人の同じアンケート(r, 全部) if answers.写真の一覧(x)],
                  key=lambda x: x.submitted_at or 昔)
    if not 写真あり:
        return {"slots": []}
    初回 = 写真あり[0]
    前 = [x for x in 写真あり if x.submitted_at and r.submitted_at and x.submitted_at < r.submitted_at]
    前回 = 前[-1] if 前 else None
    枠 = [{"label": "初回", "response": 初回}]
    if 前回 and 前回.pk != 初回.pk:
        枠.append({"label": "前回", "response": 前回})
    if answers.写真の一覧(r):
        既 = next((s for s in 枠 if s["response"].pk == r.pk), None)
        if 既:
            既["label"] = f"{既['label']} / 今回"
        else:
            枠.append({"label": "今回", "response": r})
    else:
        枠.append({"label": "今回", "response": None})
    for s in 枠:
        x = s["response"]
        if x:
            s["date"] = answers.日時の文字(x.submitted_at)
            s["name"] = 控え.会員番号付きの名前(x)
            s["ticket"] = answers.回数券情報(x)
            s["groups"] = _写真の群(x, survey, 控え)
    return {"slots": 枠}


def _編集欄(q: dict | None, a: dict) -> dict:
    """renderEditableAnswerField: 設問の種類ごとの入力欄の材料。"""
    q = q or {}
    type_ = q.get("type") or a.get("type") or "text"
    qid = a["questionId"]
    欄 = {"question_id": qid, "label": a["label"], "type": type_, "value": answers.文字(a.get("value")),
         "required": q.get("required", True) is not False, "options": q.get("options") or [], "name": f"a-{qid}"}
    if type_ == "checkbox":
        欄["selected"] = answers.回答の値一覧(a)
        欄["summary"] = None
        if qid == answers.お悩みの設問ID:
            欄["summary"] = answers.お悩みの群(a)
        elif qid == answers.変化の設問ID:
            欄["summary"] = answers.変化の群(a)
    if type_ == "rating":
        欄["options"] = ["1", "2", "3", "4", "5"]
    return 欄


def _詳細の材料(r, 全部, 控え):
    survey = r.survey or Survey.objects.filter(survey_id=r.survey_key).first()
    設問 = answers.設問の表(survey)
    表示 = answers.表示用の回答(r, survey)
    写真 = answers.写真を設問ごとに(r)
    欄 = [_編集欄(設問.get(a["questionId"]), a) for a in 表示
         if (設問.get(a["questionId"]) or {}).get("type", a.get("type")) != "photo" and not 写真.get(a["questionId"])]
    return {
        "r": r, "survey": survey, **_状況の表示(r),
        "name": 控え.会員番号付きの名前(r), "date": answers.日時の文字(r.submitted_at),
        "ticket": answers.回数券情報(r),
        "photo_groups": _写真の群(r, survey, 控え),
        "fields": 欄,
        "comparison": 前回比較(r, survey, 全部),
        "photo_comparison": 写真比較(r, survey, 全部, 控え),
        "delete_label": "完全削除" if r.trashed else "ゴミ箱へ移動",
        "delete_confirm": "この回答を完全削除しますか？" if r.trashed else "この回答をゴミ箱へ移動しますか？",
    }


@owner_required
def response_detail(request, response_id):
    r = get_object_or_404(Response.objects.prefetch_related("photos"), response_id=response_id)
    控え = answers.顧客の控え(CustomerProfile.objects.all())
    全部 = _全部()
    back = (request.GET.get("back") or "").strip()
    return render(request, "bijiris/response_detail.html", {
        **_画面の共通(),
        **_詳細の材料(r, 全部, 控え),
        "back_query": back,
        "statuses": [(k, v) for k, v in answers.対応状況の名.items()],
    })


# ---------------------------------------------------------------------------
# 書き込み（保存・ゴミ箱・完全削除）
# ---------------------------------------------------------------------------

def _断る(request, response_id=None):
    if gate.ビジリスはサーバーが正():
        return None
    messages.error(request, gate.断る文())
    if response_id:
        return redirect("manage:bijiris:response_detail", response_id=response_id)
    return redirect("manage:bijiris:response_list")


def _見える(q: dict, 生の表: dict) -> bool:
    """isQuestionVisible_ の表示条件の部分: 条件が全部一致したときだけ表示。条件が無ければ表示。
    （旧設問ごとの特別な決まり（施術回数・初回計測時の非表示など）は写していない。答えがあればそのまま残す）"""
    条件 = [c for c in (q.get("visibilityConditions") or []) if isinstance(c, dict)]
    if not 条件 and isinstance(q.get("visibleWhen"), dict):
        条件 = [q["visibleWhen"]]
    for c in 条件:
        値たち = 生の表.get(answers.文字(c.get("questionId"))) or []
        if not 値たち or answers.文字(c.get("value")) not in 値たち:
            return False
    return True


def 管理の回答を整える(survey, 既存の回答, 送られた: dict) -> list:
    """normalizeAdminAnswers_: 設問の並びで作り直す。写真はそのまま。見えない設問は空。
    5段階評価は 1〜5、選択式は選択肢の中から。"""
    既存 = {a.get("questionId"): a for a in (既存の回答 or []) if isinstance(a, dict)}
    生の表 = {qid: [v for v in 値 if v] for qid, 値 in 送られた.items()}
    出 = []
    for q in survey.questions or []:
        if not isinstance(q, dict):
            continue
        qid = q.get("id")
        元 = 既存.get(qid) or {"questionId": qid, "label": q.get("label"), "type": q.get("type"), "value": ""}
        見える = _見える(q, 生の表)
        if q.get("type") == "photo":
            出.append(元 if 見える else {**元, "value": "", "files": []})
            continue
        値たち = [answers.文字(v) for v in 生の表.get(qid, [])] if qid in 生の表 else answers.回答の値一覧(元)
        値たち = [v for v in 値たち if v]
        値 = ", ".join(値たち) if 見える else ""
        if 見える and q.get("type") == "rating" and 値 and 値 not in ("1", "2", "3", "4", "5"):
            raise ValueError("評価は1から5で回答してください。")
        if 見える and q.get("type") == "choice" and 値 and 値 not in (q.get("options") or []):
            raise ValueError("選択肢から回答してください。")
        if 見える and q.get("type") == "checkbox" and any(v not in (q.get("options") or []) for v in 値たち):
            raise ValueError("選択肢から回答してください。")
        出.append({"questionId": qid, "label": q.get("label"), "type": q.get("type"), "value": 値})
    return 出


@owner_required
@require_POST
def response_save(request, response_id):
    """updateResponse_: 対応状況・管理メモ・回答を保存し、管理更新日時を進める。"""
    if (r := _断る(request, response_id)):
        return r
    res = get_object_or_404(Response, response_id=response_id)
    survey = res.survey or Survey.objects.filter(survey_id=res.survey_key).first()
    送られた = {k[2:]: request.POST.getlist(k) for k in request.POST if k.startswith("a-")}
    try:
        回答 = 管理の回答を整える(survey, res.answers, 送られた) if survey else res.answers
    except ValueError as e:
        messages.error(request, str(e))
        return redirect("manage:bijiris:response_detail", response_id=response_id)
    res.status = answers.対応状況を整える(request.POST.get("status"))
    res.admin_memo = (request.POST.get("admin_memo") or "").strip()
    res.answers = 回答
    res.managed_at = timezone.now()
    res.save()
    messages.success(request, "回答管理を保存しました。")
    return redirect(_詳細へ(request, response_id))


def _詳細へ(request, response_id) -> str:
    back = (request.POST.get("back") or "").strip()
    url = f"/manage/bijiris/responses/{response_id}/"
    return f"{url}?{urlencode({'back': back})}" if back else url


@owner_required
@require_POST
def response_trash(request, response_id):
    """trashResponse_: 対応状況を「ゴミ箱」に。"""
    if (r := _断る(request, response_id)):
        return r
    res = get_object_or_404(Response, response_id=response_id)
    res.status = "trash"
    res.managed_at = timezone.now()
    res.save(update_fields=["status", "managed_at", "changed_at"])
    messages.success(request, "回答をゴミ箱へ移動しました。")
    back = (request.POST.get("back") or "").strip()
    return redirect(f"/manage/bijiris/responses/?{back}" if back else "manage:bijiris:response_list")


def _写真のファイルを消す(res):
    """media/bijiris/ に入れた写真の実体を消す（purgeResponse_ の deleteResponseFiles_ にあたる）。"""
    for p in res.photos.all():
        if not p.url:
            continue
        名前 = p.url.rsplit("/", 1)[-1]
        f = Path(settings.MEDIA_ROOT) / "bijiris" / 名前
        if 名前 and f.is_file():
            f.unlink()


@owner_required
@require_POST
def response_purge(request, response_id):
    """完全削除: ゴミ箱の回答だけ、写真ごと消す。"""
    if (r := _断る(request, response_id)):
        return r
    res = get_object_or_404(Response, response_id=response_id)
    if not res.trashed:
        messages.error(request, "ゴミ箱に入っていない回答は完全削除できません。先にゴミ箱へ移動してください。")
        return redirect("manage:bijiris:response_detail", response_id=response_id)
    _写真のファイルを消す(res)
    res.delete()
    messages.success(request, "回答を完全削除しました。")
    back = (request.POST.get("back") or "").strip()
    return redirect(f"/manage/bijiris/responses/?{back}" if back else "manage:bijiris:response_list")


# ---------------------------------------------------------------------------
# CSV・印刷
# ---------------------------------------------------------------------------

def _行(r, 控え) -> list[str]:
    券 = dict(answers.回数券情報(r))
    return [
        answers.日時の文字(r.submitted_at),
        控え.会員番号(r),
        r.customer_name or "",
        r.survey_title or "",
        券.get("回数券", ""), 券.get("何枚目", ""), 券.get("何回目", ""),
        answers.対応状況の名[answers.対応状況を整える(r.status)],
        r.admin_memo or "",
        str(len(answers.写真の一覧(r))),
        " / ".join(answers.回答文の一覧(r)),
    ]


@owner_required
def response_export(request):
    """exportCsv: 絞り込んだ回答を CSV（BOM 付き UTF-8）で。"""
    値 = _絞り込みの値(request)
    控え = answers.顧客の控え(CustomerProfile.objects.all())
    絞った = 絞り込む(_全部(), 値, 控え)
    if not 絞った:
        messages.error(request, "出力対象の回答がありません。")
        return redirect(f"/manage/bijiris/responses/?{_絞り込みの文字列(値)}")
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="\n")
    w.writerow(["日時", "会員番号", "お客様", "アンケート", "回数券", "何枚目", "何回目", "対応状況", "管理メモ", "写真枚数", "回答"])
    for r in 絞った:
        w.writerow(_行(r, 控え))
    # 先頭の BOM は Excel で開いたときに文字化けしないため（旧アプリも付けている）
    res = HttpResponse("\ufeff" + buf.getvalue(), content_type="text/csv; charset=utf-8")
    res["Content-Disposition"] = f'attachment; filename="survey-responses-{int(timezone.now().timestamp() * 1000)}.csv"'
    return res


@owner_required
def response_print_list(request):
    """exportResponsesPdf: 絞り込んだ回答を印刷用の1枚に（ブラウザの印刷で PDF に）。"""
    値 = _絞り込みの値(request)
    控え = answers.顧客の控え(CustomerProfile.objects.all())
    絞った = 絞り込む(_全部(), 値, 控え)
    if not 絞った:
        messages.error(request, "出力対象の回答がありません。")
        return redirect(f"/manage/bijiris/responses/?{_絞り込みの文字列(値)}")
    項 = [{
        "r": r, "date": answers.日時の文字(r.submitted_at), "name": 控え.会員番号付きの名前(r),
        "status_label": answers.対応状況の名[answers.対応状況を整える(r.status)],
        "ticket": " / ".join(f"{k} {v}" for k, v in answers.回数券情報(r)) or "-",
        "photo_count": len(answers.写真の一覧(r)), "answers_text": " / ".join(answers.回答文の一覧(r)),
    } for r in 絞った]
    return render(request, "bijiris/response_print.html", {"mode": "list", "items": 項, "count": len(項)})


@owner_required
def response_print(request, response_id):
    """printResponse: 1件を印刷用に。"""
    r = get_object_or_404(Response.objects.prefetch_related("photos"), response_id=response_id)
    控え = answers.顧客の控え(CustomerProfile.objects.all())
    survey = r.survey or Survey.objects.filter(survey_id=r.survey_key).first()
    写真 = answers.写真を設問ごとに(r)
    項 = []
    for a in answers.表示用の回答(r, survey):
        ps = 写真.get(a["questionId"])
        項.append({"label": a["label"], "value": ", ".join((p.name or "写真") for p in ps) if ps else (answers.文字(a.get("value")) or "未回答")})
    return render(request, "bijiris/response_print.html", {
        "mode": "single", "r": r, "name": 控え.会員番号付きの名前(r), "date": answers.日時の文字(r.submitted_at), "answers": 項,
    })
