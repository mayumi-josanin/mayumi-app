"""ビジリス管理「集計」。**中身は旧ビジリス管理アプリ（#page-dashboard、app.js renderDashboard 1922行〜）と同じ。**

出すもの（上から順に）:
  1. 回答数・未対応・回答者数・アンケート数（statsGrid）
  2. 最新の回答 5件（latestResponses）
  3. アンケート別回答数（surveySummary）
  4. 質問別分析（renderSurveyAnalytics / getSurveyAnalyticsSummary。アンケートを選ぶ）
  5. 月別回答件数（renderMonthlySurveyChart / getMonthlySurveyCounts。直近6か月）

数え方は旧アプリと同じ: **ゴミ箱（status=trash）は数えない。**アンケートは全部（下書き・アーカイブも）数える。
読むだけの画面なので切り替えの印は見ない。
"""

from collections import Counter
from datetime import datetime
from datetime import timezone as dt_timezone

from django.shortcuts import render
from django.utils import timezone

from apps.manage.permissions import owner_required

from . import answers
from .models import CustomerProfile, Response, Survey

# 送信日時が空の回答（取り込みで読めなかったもの）を並べるときの仮の値。一番古い扱い
昔 = datetime(2000, 1, 1, tzinfo=dt_timezone.utc)


def _有効な回答():
    """renderDashboard の activeResponses: ゴミ箱を除いた全部。"""
    return list(Response.objects.exclude(status="trash").prefetch_related("photos").order_by("-submitted_at", "-sheet_row"))


def 質問別分析(survey, responses) -> list[dict]:
    """getSurveyAnalyticsSummary と同じ。設問ごとに
    - 写真: 「写真添付あり」の件数
    - 単一選択・複数選択・5段階評価: 値ごとの件数（多い順）。無ければ「未回答 0件」
      （お悩みの設問だけは分類ごとにまとめる: getConcernAnalyticsRows）
    - それ以外: 「入力あり」の件数
    """
    回答の表たち = [answers.回答の表(r) for r in responses]
    写真の表たち = [answers.写真を設問ごとに(r) for r in responses]
    出 = []
    for q in survey.questions or []:
        if not isinstance(q, dict):
            continue
        qid = q.get("id")
        項 = {"question": q, "label": q.get("label") or "", "rows": [], "groups": None}
        if q.get("type") == "photo":
            # 旧アプリは answer.files の有無。ここでは取り込んだ写真の記録（ResponsePhoto）で見る
            n = sum(1 for 表, 写真 in zip(回答の表たち, 写真の表たち)
                    if 写真.get(qid) or (表.get(qid) or {}).get("files"))
            項["rows"] = [{"label": "写真添付あり", "count": n}]
        elif q.get("type") in ("checkbox", "choice", "rating"):
            if qid == answers.お悩みの設問ID:
                選択たち = [set(answers.回答の値一覧(表.get(qid))) for 表 in 回答の表たち]
                項["groups"] = [{
                    "label": c["label"],
                    "count": sum(1 for s in 選択たち if any(o in s for o in c["options"])),
                    "rows": [{"label": o, "count": sum(1 for s in 選択たち if o in s)} for o in c["options"]],
                } for c in answers.お悩みの分類]
            else:
                数 = Counter()
                for 表 in 回答の表たち:
                    a = 表.get(qid)
                    値 = answers.回答の値一覧(a) if q.get("type") == "checkbox" else [v for v in [answers.文字((a or {}).get("value"))] if v]
                    数.update(値)
                行 = [{"label": k, "count": n} for k, n in sorted(数.items(), key=lambda kv: -kv[1])]
                項["rows"] = 行 or [{"label": "未回答", "count": 0}]
        else:
            n = sum(1 for 表 in 回答の表たち if answers.文字((表.get(qid) or {}).get("value")))
            項["rows"] = [{"label": "入力あり", "count": n}]
        出.append(項)
    return 出


def 月別回答件数(responses, surveys) -> list[dict]:
    """getMonthlySurveyCounts + renderMonthlySurveyChart: 月（日本時間）× アンケート名の件数。直近6か月ぶん。
    棒の長さは全体の最大値に対する割合。"""
    月ごと = {}
    for r in responses:
        if not r.submitted_at:
            continue
        月 = timezone.localtime(r.submitted_at).strftime("%Y-%m")
        月ごと.setdefault(月, Counter())[r.survey_title] += 1
    月々 = sorted(月ごと.items())[-6:]
    題 = [s.title for s in surveys]
    最大 = max([1] + [c.get(t, 0) for _, c in 月々 for t in 題])
    return [{
        "month": 月,
        "bars": [{"title": t, "count": c.get(t, 0), "width": round(c.get(t, 0) / 最大 * 100)} for t in 題],
    } for 月, c in 月々]


@owner_required
def dashboard(request):
    surveys = list(Survey.objects.all())
    有効 = _有効な回答()
    控え = answers.顧客の控え(CustomerProfile.objects.all())

    stats = [
        ("回答数", len(有効)),
        ("未対応", sum(1 for r in 有効 if answers.対応状況を整える(r.status) == "new")),
        ("回答者数", len({r.customer_name for r in 有効})),
        ("アンケート数", len(surveys)),
    ]
    最新 = [{
        "name": 控え.会員番号付きの名前(r), "survey_title": r.survey_title, "submitted": answers.日時の文字(r.submitted_at),
        "status": answers.対応状況を整える(r.status),
        "status_label": answers.対応状況の名[answers.対応状況を整える(r.status)],
        "badge": answers.対応状況の印[answers.対応状況を整える(r.status)],
    } for r in sorted(有効, key=lambda r: r.submitted_at or 昔, reverse=True)[:5]]
    # アンケート別回答数: 回答の surveyId（survey_key）がアンケートの id と同じもの
    別 = [{"title": s.title, "count": sum(1 for r in 有効 if r.survey_key == s.survey_id)} for s in surveys]

    # 質問別分析: ?survey=<id>。無ければ先頭のアンケート
    選んだID = request.GET.get("survey") or (surveys[0].survey_id if surveys else "")
    選んだ = next((s for s in surveys if s.survey_id == 選んだID), None)
    分析 = 質問別分析(選んだ, [r for r in 有効 if r.survey_key == 選んだ.survey_id]) if 選んだ else []

    return render(request, "bijiris/dashboard.html", {
        "title": "集計",
        "stats": stats,
        "latest": 最新,
        "per_survey": 別,
        "surveys": surveys,
        "selected_survey": 選んだ,
        "analytics": 分析,
        "monthly": 月別回答件数(有効, surveys),
    })
