"""開発管理の計算（KEM の apps/devkanri/services.py の写し）。

ガントチャート（frappe-gantt）に渡す行と、進捗の計算。
KEM にあった Discord 送信はこちらでは notifications.py にだけ持つ（二重に持たない）。
"""

import datetime

from django.urls import reverse

# 日付の材料が何も無いプロジェクトのバーの長さ（日）。
# frappe-gantt は終了日に時刻が無いと24時間を足すので0日でも描画自体はされるが、
# 1日の点では「いつからいつまでか分かっていない」ことが伝わらないため幅を持たせる。
DEFAULT_SPAN_DAYS = 7

DONE_STATUSES = ("done", "closed")


def resolve_project_period(project, task_due_dates=None):
    """プロジェクトの表示期間を決める。

    開始日・期限は任意入力で、実際には未設定のプロジェクトが多い。
    そこだけガントから消えると全体像が見えないので、分かる材料で補う。

      開始日 … start_date → タスク期限の最小 → 作成日
      期限   … due_date  → タスク期限の最大 → 開始日 + DEFAULT_SPAN_DAYS

    Returns:
        (start: date, end: date, inferred: bool)
        inferred は補完した日付を含むか。画面では「推定」と分かるように出す。
    """
    dues = sorted(d for d in (task_due_dates or []) if d)

    start = project.start_date
    if not start:
        start = dues[0] if dues else _created_date(project)

    end = project.due_date
    if not end:
        end = dues[-1] if dues else None
    if not end or end < start:
        end = start + datetime.timedelta(days=DEFAULT_SPAN_DAYS)

    inferred = not (project.start_date and project.due_date)
    return start, end, inferred


def _created_date(project):
    """作成日時から日付を取り出す。created_at が無い場合は本日。"""
    created = getattr(project, "created_at", None)
    if not created:
        return datetime.date.today()
    return created.date() if hasattr(created, "date") else created


def calc_progress(project, tasks) -> int:
    """タスクの完了割合（％）。タスクが無ければステータスから決める。"""
    total = len(tasks)
    if total == 0:
        return 100 if project.status == "completed" else 0
    done = sum(1 for t in tasks if t.status in DONE_STATUSES)
    return int(done / total * 100)


def get_project_gantt_data(projects) -> list[dict]:
    """開発プロジェクトのガントチャートデータ（frappe-gantt用）。

    projects は tasks を prefetch しておくとクエリが1回で済む。返すのは期間の早い順。
    """
    rows = []
    for project in projects:
        tasks = list(project.tasks.all())
        start, end, inferred = resolve_project_period(project, [t.due_date for t in tasks])
        css = [f"bar-dev-{project.status}"]
        if inferred:
            css.append("bar-dev-inferred")

        rows.append({
            "id": f"project-{project.pk}",
            "name": project.name,
            "start": start.isoformat(),
            "end": end.isoformat(),
            "progress": calc_progress(project, tasks),
            "custom_class": " ".join(css),
            # 以下はポップアップ表示用。frappe-gantt は未知のキーを保持する。
            "assignee": str(project.assignee) if project.assignee else "未割当",
            "status_label": project.get_status_display(),
            "inferred": inferred,
            "task_total": len(tasks),
            "task_done": sum(1 for t in tasks if t.status in DONE_STATUSES),
            "url": reverse("manage:dev:project_detail", args=[project.pk]),
        })

    rows.sort(key=lambda r: (r["start"], r["end"]))
    return rows
