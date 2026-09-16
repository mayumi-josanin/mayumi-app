"""ビジリス管理「アンケート管理」。**中身は旧ビジリス管理アプリ（#page-surveys、app.js renderSurveyManager 7883行〜）と同じ。**

一覧（survey_list）: 「N. タイトル」・公開状態・設問数・更新日時・受付中/受付停止中、↑ ↓ 複製 アーカイブ。
編集（survey_form）: タイトル・説明・回答前メッセージ・回答完了メッセージ・公開状態・回答受付中・受付開始/終了日時・
設問（質問文・質問形式・必須・表示条件・選択肢）。設問の並べ替えと入力は画面の JS（旧アプリと同じ入力欄）で行い、
保存のときだけ JSON にまとめて送る（院長が JSON を直に触ることはない）。

保存の規則は Code.gs の validateSurveyPayload_（865行〜）・createSurvey_（973行〜）・updateSurveyDefinition_・
deleteSurveyDefinition_・replaceSurveys_（並び順）・normalizeSurveyOrder_ と同じ。文言もそのまま。

書くのは切り替えの印（gate）が立っているときだけ。立つまでは断って一覧へ戻す。
"""

import json
import uuid

from django.contrib import messages
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_datetime
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import answers, gate
from .models import Survey

公開状態の名 = {"published": "公開中", "draft": "下書き", "archived": "アーカイブ"}
公開状態の印 = {"published": "badge-green", "draft": "badge-gray", "archived": "badge-orange"}
回答完了の既定文 = "ご回答ありがとうございました。"


def 公開状態を整える(status) -> str:
    return status if status in 公開状態の名 else "published"


def _文字(v) -> str:
    return answers.文字(v)


def 日時を整える(値):
    """normalizeDateTime_: 空なら None。読めなければ「日時の形式が正しくありません。」。
    画面の datetime-local（時刻帯なし）は日本時間として受ける。"""
    s = _文字(値)
    if not s:
        return None
    dt = parse_datetime(s)
    if dt is None:
        raise ValueError("日時の形式が正しくありません。")
    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt, timezone.get_current_timezone())
    return dt


def 表示条件を整える(条件たち, visible_when) -> list[dict]:
    """validateVisibilityConditions_: {questionId, value} の両方が入っているものだけ。無ければ visibleWhen を1つ。"""
    出 = []
    for c in 条件たち if isinstance(条件たち, list) else []:
        if isinstance(c, dict) and _文字(c.get("questionId")) and _文字(c.get("value")):
            出.append({"questionId": _文字(c.get("questionId")), "value": _文字(c.get("value"))})
    if 出:
        return 出
    if isinstance(visible_when, dict) and _文字(visible_when.get("questionId")) and _文字(visible_when.get("value")):
        return [{"questionId": _文字(visible_when.get("questionId")), "value": _文字(visible_when.get("value"))}]
    return []


def アンケートを整える(payload: dict, existing: Survey | None) -> dict:
    """validateSurveyPayload_ と同じ検査と既定値。返すのは Survey の項目の dict。"""
    title = _文字(payload.get("title"))
    description = _文字(payload.get("description"))
    intro = _文字(payload.get("introMessage")) or description
    completion = _文字(payload.get("completionMessage")) or 回答完了の既定文
    status = 公開状態を整える(_文字(payload.get("status")))
    questions = payload.get("questions") if isinstance(payload.get("questions"), list) else []
    accepting = payload.get("acceptingResponses") is not False
    start_at = 日時を整える(payload.get("startAt"))
    end_at = 日時を整える(payload.get("endAt"))

    if not title:
        raise ValueError("タイトルを入力してください。")
    if not description:
        raise ValueError("説明文を入力してください。")
    if not questions:
        raise ValueError("質問は1つ以上必要です。")
    if start_at and end_at and start_at > end_at:
        raise ValueError("受付終了日時は開始日時以降にしてください。")

    設問 = []
    for q in questions:
        q = q if isinstance(q, dict) else {}
        type_ = q.get("type") if q.get("type") in answers.設問の種類 else "text"
        label = _文字(q.get("label"))
        options = [_文字(o) for o in (q.get("options") if isinstance(q.get("options"), list) else [])]
        options = [o for o in options if o]
        条件 = 表示条件を整える(q.get("visibilityConditions"), q.get("visibleWhen"))
        if not label:
            raise ValueError("質問文を入力してください。")
        選択式 = type_ in ("choice", "checkbox")
        if 選択式 and len(options) < 2:
            raise ValueError("選択式の質問は選択肢を2つ以上入力してください。")
        設問.append({
            "id": _文字(q.get("id")) or f"question_{uuid.uuid4()}",
            "label": label,
            "type": type_,
            "required": q.get("required") is not False,
            "options": options if 選択式 else [],
            "placeholder": _文字(q.get("placeholder")),
            "visibilityConditions": 条件,
            "visibleWhen": 条件[0] if 条件 else None,
        })

    now = timezone.now()
    return {
        "title": title, "description": description, "intro_message": intro, "completion_message": completion,
        "status": status, "accepting_responses": accepting, "start_at": start_at, "end_at": end_at,
        "questions": 設問,
        "created_at": existing.created_at if existing and existing.created_at else now,
        "updated_at": now,
    }


def 並びを整える():
    """normalizeSurveyOrder_: いまの並び（sort_order, survey_id）で 0 から振り直す。"""
    for i, s in enumerate(Survey.objects.order_by("sort_order", "survey_id")):
        if s.sort_order != i:
            Survey.objects.filter(pk=s.pk).update(sort_order=i)


def _断る(request):
    """印が立っていなければ断って一覧へ。立っていれば None。"""
    if gate.ビジリスはサーバーが正():
        return None
    messages.error(request, gate.断る文())
    return redirect("manage:bijiris:survey_list")


def _画面の共通():
    return {"title": "アンケート管理", "server_is_source": gate.ビジリスはサーバーが正(), "refuse_text": gate.断る文()}


def _一覧の1件(s: Survey, i: int) -> dict:
    return {
        "s": s, "no": i + 1, "status": 公開状態を整える(s.status),
        "status_label": 公開状態の名[公開状態を整える(s.status)], "badge": 公開状態の印[公開状態を整える(s.status)],
        "updated": answers.日時の文字(s.updated_at), "question_count": s.question_count,
        "accepting": "受付中" if s.accepting_responses else "受付停止中",
    }


@owner_required
def survey_list(request):
    surveys = list(Survey.objects.all())
    return render(request, "bijiris/survey_list.html", {
        **_画面の共通(),
        "items": [_一覧の1件(s, i) for i, s in enumerate(surveys)],
        "count": len(surveys),
    })


# ---------------------------------------------------------------------------
# 編集画面
# ---------------------------------------------------------------------------

def _設問を編集用に(q: dict) -> dict:
    """normalizeQuestionForEditor: 画面の JS に渡す形。"""
    q = q if isinstance(q, dict) else {}
    return {
        "id": q.get("id") or f"question_{uuid.uuid4().hex[:12]}",
        "label": q.get("label") or "",
        "type": q.get("type") or "text",
        "required": q.get("required") is not False,
        "options": q.get("options") if isinstance(q.get("options"), list) else [],
        "visibilityConditions": 表示条件を整える(q.get("visibilityConditions"), q.get("visibleWhen")),
    }


def _空の設問() -> dict:
    return _設問を編集用に({})


def _局所の日時(dt) -> str:
    """datetime-local の value（日本時間）。toDateTimeLocalValue と同じ。"""
    return timezone.localtime(dt).strftime("%Y-%m-%dT%H:%M") if dt else ""


def _下書き(s: Survey | None) -> dict:
    """getSurveyEditorDraft: 保存されている内容を画面の値に。新規は空（設問1つ）。"""
    if not s:
        return {"title": "", "description": "", "introMessage": "", "completionMessage": "", "status": "published",
                "acceptingResponses": True, "startAt": "", "endAt": "", "questions": [_空の設問()]}
    設問 = [_設問を編集用に(q) for q in (s.questions or [])]
    return {
        "title": s.title, "description": s.description, "introMessage": s.intro_message,
        "completionMessage": s.completion_message, "status": 公開状態を整える(s.status),
        "acceptingResponses": s.accepting_responses, "startAt": _局所の日時(s.start_at), "endAt": _局所の日時(s.end_at),
        "questions": 設問 or [_空の設問()],
    }


def _送られた下書き(request) -> dict:
    """POST の値を、保存に使う形（validateSurveyPayload_ の payload）に。設問は JSON（画面の JS がまとめる）。"""
    try:
        設問 = json.loads(request.POST.get("questions") or "[]")
    except ValueError:
        設問 = []
    return {
        "title": request.POST.get("title", ""),
        "description": request.POST.get("description", ""),
        "introMessage": request.POST.get("introMessage", ""),
        "completionMessage": request.POST.get("completionMessage", ""),
        "status": request.POST.get("status", "published"),
        "acceptingResponses": bool(request.POST.get("acceptingResponses")),
        "startAt": request.POST.get("startAt", ""),
        "endAt": request.POST.get("endAt", ""),
        "questions": 設問 if isinstance(設問, list) else [],
    }


def _編集画面(request, s: Survey | None, draft: dict):
    draft = dict(draft)
    draft["questions"] = [_設問を編集用に(q) for q in draft.get("questions") or []] or [_空の設問()]
    return render(request, "bijiris/survey_form.html", {
        **_画面の共通(),
        "survey": s,
        "draft": draft,
        "form_title": "アンケート編集" if s else "アンケート新規作成",
        "created": answers.日時の文字(s.created_at) if s else "",
        "updated": answers.日時の文字(s.updated_at) if s else "",
        "concern_categories": answers.お悩みの分類,
        "concern_question_id": answers.お悩みの設問ID,
        "question_types": [(k, answers.設問の種類の名[k]) for k in answers.設問の種類],
    })


def _保存の文言(新規: bool, status: str) -> str:
    公開 = 公開状態を整える(status) == "published"
    if 新規:
        return "アンケートを公開しました。" if 公開 else "アンケートを作成しました。"
    return "アンケートの公開内容を保存しました。" if 公開 else "アンケートを保存しました。"


@owner_required
def survey_create(request):
    if request.method == "POST":
        if (r := _断る(request)):
            return r
        payload = _送られた下書き(request)
        try:
            値 = アンケートを整える(payload, None)
        except ValueError as e:
            messages.error(request, str(e))
            return _編集画面(request, None, payload)
        with transaction.atomic():
            # createSurvey_: 新しいものは並びの最後（sortOrder = 本数）に付き、そのあと 0 から振り直す
            s = Survey.objects.create(survey_id=f"survey_{uuid.uuid4()}", sort_order=Survey.objects.count(), **値)
            並びを整える()
        messages.success(request, _保存の文言(True, s.status))
        return redirect("manage:bijiris:survey_edit", survey_id=s.survey_id)
    return _編集画面(request, None, _下書き(None))


@owner_required
def survey_edit(request, survey_id):
    s = get_object_or_404(Survey, survey_id=survey_id)
    if request.method == "POST":
        if (r := _断る(request)):
            return r
        payload = _送られた下書き(request)
        try:
            値 = アンケートを整える(payload, s)
        except ValueError as e:
            messages.error(request, str(e))
            return _編集画面(request, s, payload)
        with transaction.atomic():
            for k, v in 値.items():
                setattr(s, k, v)
            s.save()
            並びを整える()
        messages.success(request, _保存の文言(False, s.status))
        return redirect("manage:bijiris:survey_edit", survey_id=s.survey_id)
    return _編集画面(request, s, _下書き(s))


# ---------------------------------------------------------------------------
# 一覧の操作（↑ ↓ 複製 アーカイブ 削除）
# ---------------------------------------------------------------------------

@owner_required
@require_POST
def survey_move(request, survey_id):
    """moveSurvey → replaceSurveys_: 隣と入れ替えて 0 から振り直す。端では何もしない。"""
    if (r := _断る(request)):
        return r
    s = get_object_or_404(Survey, survey_id=survey_id)
    向き = -1 if request.POST.get("direction") == "-1" else 1
    with transaction.atomic():
        並びを整える()
        並び = list(Survey.objects.order_by("sort_order", "survey_id"))
        i = next(k for k, x in enumerate(並び) if x.pk == s.pk)
        j = i + 向き
        if 0 <= j < len(並び):
            Survey.objects.filter(pk=並び[i].pk).update(sort_order=j)
            Survey.objects.filter(pk=並び[j].pk).update(sort_order=i)
    return redirect("manage:bijiris:survey_list")


@owner_required
@require_POST
def survey_duplicate(request, survey_id):
    """duplicateSurvey: 「（複製）」を付けた下書きとして新しく作る（createSurvey_ と同じ）。"""
    if (r := _断る(request)):
        return r
    s = get_object_or_404(Survey, survey_id=survey_id)
    payload = {
        "title": f"{s.title}（複製）", "description": s.description, "introMessage": s.intro_message,
        "completionMessage": s.completion_message, "status": "draft", "acceptingResponses": s.accepting_responses,
        "startAt": s.start_at.isoformat() if s.start_at else "", "endAt": s.end_at.isoformat() if s.end_at else "",
        "questions": json.loads(json.dumps(s.questions or [])),
    }
    try:
        値 = アンケートを整える(payload, None)
    except ValueError as e:
        messages.error(request, str(e) or "アンケートを複製できませんでした。")
        return redirect("manage:bijiris:survey_list")
    with transaction.atomic():
        新 = Survey.objects.create(survey_id=f"survey_{uuid.uuid4()}", sort_order=Survey.objects.count(), **値)
        並びを整える()
    messages.success(request, "アンケートを複製しました。")
    return redirect("manage:bijiris:survey_edit", survey_id=新.survey_id)


@owner_required
@require_POST
def survey_archive(request, survey_id):
    """archiveSurvey: 公開状態だけ「アーカイブ」に（ほかは保存済みのまま。updatedAt は進む）。"""
    if (r := _断る(request)):
        return r
    s = get_object_or_404(Survey, survey_id=survey_id)
    s.status = "archived"
    s.updated_at = timezone.now()
    s.save(update_fields=["status", "updated_at", "changed_at"])
    messages.success(request, "アンケートをアーカイブしました。")
    return redirect(request.POST.get("next") or "manage:bijiris:survey_list")


@owner_required
@require_POST
def survey_delete(request, survey_id):
    """deleteSurveyDefinition_: 定義を消す。回答は残る（survey が空になるだけ。GAS もシートは消さない）。"""
    if (r := _断る(request)):
        return r
    s = get_object_or_404(Survey, survey_id=survey_id)
    with transaction.atomic():
        s.delete()
        並びを整える()
    messages.success(request, "アンケートを削除しました。")
    return redirect("manage:bijiris:survey_list")
