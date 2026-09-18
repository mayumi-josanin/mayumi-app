"""開発管理の画面（KEM の apps/devkanri/views.py の写し）。

全部 **まゆみだけ**（owner_required）。スタッフには左メニューにも出るが 403。
KEM は login_required ＋ 会社（tenant）で絞っていた。こちらは会社が無いので絞りは要らない。
"""

import datetime

from django.contrib import messages
from django.db.models import Case, Count, F, IntegerField, Q, Sum, When
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.http import url_has_allowed_host_and_scheme

from apps.manage import images
from apps.manage.permissions import owner_required

from .forms import DevCommentForm, DevProjectForm, DevTaskForm, MeyasubakoForm
from .models import DevProject, DevTask, Meyasubako
from .notifications import notify_task_assigned, notify_task_status_changed
from .services import get_project_gantt_data

# 一覧の絞り込み。既定は動いているものだけを出す。
# 完了・中断まで並べるとガントが埋まって「いま何が動いているか」が読めなくなる。
PROJECT_SCOPES = {
    "active": "稼働中",
    "mine": "自分の担当",
    "all": "すべて",
}
ACTIVE_STATUSES = ["planning", "in_progress"]

# タスク一覧（プロジェクトをまたいだ一覧。院長の依頼 2026-09-18）。
# 「いま何が残っているか」を見る画面なので、**既定では完了・クローズを隠す**。
# プロジェクト詳細だけだと、残りを見るのにプロジェクトを1つずつ開くことになっていた。
DONE_STATUSES = ["done", "closed"]
TASK_SORTS = {
    "priority": "優先度の高い順",
    "due": "期限の近い順",
    "new": "新しい順",
}
# 優先度は文字なので、そのまま並べると英字順（critical→high→low→medium）になってしまう。
# 数に置き換えてから並べる。
PRIORITY_ORDER = ["critical", "high", "medium", "low"]


@owner_required
def project_list(request):
    """プロジェクト一覧。1ページ目はガントチャートで全体像を見せる。"""
    scope = request.GET.get("scope", "active")
    if scope not in PROJECT_SCOPES:
        scope = "active"

    projects = DevProject.objects.annotate(
        task_count=Count("tasks"),
        done_count=Count("tasks", filter=Q(tasks__status__in=["done", "closed"])),
        bug_count=Count("tasks", filter=Q(tasks__category="bug") & ~Q(tasks__status__in=["done", "closed"])),
    ).select_related("assignee").prefetch_related("tasks")

    if scope == "active":
        projects = projects.filter(status__in=ACTIVE_STATUSES)
    elif scope == "mine":
        projects = projects.filter(assignee=request.user)

    projects = projects.order_by("-created_at")

    gantt_rows = get_project_gantt_data(projects)

    # 表にも同じ期間を出せるよう、行をプロジェクトに持たせる。
    # 並びもガント（期間の早い順）に合わせて、目で追えるようにする。
    row_by_id = {row["id"]: row for row in gantt_rows}
    ordered = []
    for project in projects:
        project.period = row_by_id.get(f"project-{project.pk}")
        ordered.append(project)
    ordered.sort(key=lambda p: (p.period["start"], p.period["end"]) if p.period else ("", ""))

    return render(request, "dev/project_list.html", {
        "projects": ordered,
        # テンプレートでは json_script で埋める（プロジェクト名に "</script>" が入っても壊れない）
        "gantt_rows": gantt_rows,
        "gantt_tasks_exist": len(gantt_rows) > 0,
        "scope": scope,
        "scopes": PROJECT_SCOPES,
        "inferred_count": sum(1 for row in gantt_rows if row["inferred"]),
    })


@owner_required
def task_list(request):
    """タスク一覧。プロジェクトをまたいで、残っているタスクをまとめて出す。

    絞り込みと並べ替えは住所（クエリ文字列）に持つ。行からステータスを進めたときも
    同じ住所に戻すので、絞り込んだ状態のまま次の行に進める。
    """
    today = timezone.localdate()
    # 今週の終わり（月曜はじまりの日曜まで）。「今週が期限」の札に使う
    週末 = today + datetime.timedelta(days=6 - today.weekday())

    show_done = request.GET.get("done") == "1"
    overdue_only = request.GET.get("overdue") == "1"
    sort = request.GET.get("sort") or "priority"
    if sort not in TASK_SORTS:
        sort = "priority"

    条件 = {k: (request.GET.get(k) or "") for k in ("project", "status", "priority", "category", "assignee")}

    tasks = DevTask.objects.select_related("project", "assignee")
    if not show_done:
        tasks = tasks.exclude(status__in=DONE_STATUSES)
    if 条件["project"].isdigit():
        tasks = tasks.filter(project_id=int(条件["project"]))
    if 条件["status"] in dict(DevTask.Status.choices):
        tasks = tasks.filter(status=条件["status"])
    else:
        条件["status"] = ""
    if 条件["priority"] in dict(DevTask.Priority.choices):
        tasks = tasks.filter(priority=条件["priority"])
    else:
        条件["priority"] = ""
    if 条件["category"] in dict(DevTask.Category.choices):
        tasks = tasks.filter(category=条件["category"])
    else:
        条件["category"] = ""
    if 条件["assignee"].isdigit():
        tasks = tasks.filter(assignee_id=int(条件["assignee"]))
    if overdue_only:
        # 終わったものは、期限を過ぎていても「期限切れ」に数えない
        tasks = tasks.filter(due_date__lt=today).exclude(status__in=DONE_STATUSES)

    tasks = tasks.annotate(優先度順=Case(
        *[When(priority=v, then=i) for i, v in enumerate(PRIORITY_ORDER)],
        default=len(PRIORITY_ORDER), output_field=IntegerField(),
    ))
    if sort == "due":
        # 期限なしは後ろへ（期限の近い順で見たいのは、期限のあるもの）
        tasks = tasks.order_by(F("due_date").asc(nulls_last=True), "優先度順", "-created_at")
    elif sort == "new":
        tasks = tasks.order_by("-created_at")
    else:
        tasks = tasks.order_by("優先度順", F("due_date").asc(nulls_last=True), "-created_at")

    # プロジェクトごとにまとめる。**並びは最初に出てきた順**にして、
    # 並べ替え（優先度の高い順など）が見出しの順にもそのまま出るようにする
    groups = {}
    集計 = {"remaining": 0, "in_progress": 0, "overdue": 0, "this_week": 0}
    for t in tasks:
        終わった = t.status in DONE_STATUSES
        t.is_overdue = bool(t.due_date and t.due_date < today and not 終わった)
        t.is_soon = bool(t.due_date and today <= t.due_date <= today + datetime.timedelta(days=1) and not 終わった)
        g = groups.setdefault(t.project_id, {"project": t.project, "tasks": [], "remaining": 0, "overdue": 0})
        g["tasks"].append(t)
        if not 終わった:
            g["remaining"] += 1
            集計["remaining"] += 1
            if t.status == "in_progress":
                集計["in_progress"] += 1
            if t.due_date and today <= t.due_date <= 週末:
                集計["this_week"] += 1
        if t.is_overdue:
            g["overdue"] += 1
            集計["overdue"] += 1

    # 絞り込みの選択肢。担当は、実際にタスクを持っている人だけ出す（使わない名前を並べない）
    from django.contrib.auth import get_user_model

    担当たち = get_user_model().objects.filter(dev_tasks__isnull=False).distinct().order_by("username")

    return render(request, "dev/task_list.html", {
        "groups": list(groups.values()),
        "count": sum(len(g["tasks"]) for g in groups.values()),
        "summary": 集計,
        "sort": sort,
        "sorts": TASK_SORTS,
        "show_done": show_done,
        "overdue_only": overdue_only,
        "filters": 条件,
        "projects": DevProject.objects.order_by("name"),
        "assignees": 担当たち,
        "status_choices": DevTask.Status.choices,
        "priority_choices": DevTask.Priority.choices,
        "category_choices": DevTask.Category.choices,
        "here": request.get_full_path(),
    })


@owner_required
def project_detail(request, pk):
    project = get_object_or_404(DevProject, pk=pk)
    tasks = project.tasks.select_related("assignee").all()

    # 集計
    total = tasks.count()
    done = tasks.filter(status__in=["done", "closed"]).count()
    in_progress = tasks.filter(status="in_progress").count()
    open_count = tasks.filter(status="open").count()
    review_count = tasks.filter(status="review").count()
    open_bugs = tasks.filter(category="bug").exclude(status__in=["done", "closed"]).count()
    hours = tasks.aggregate(est=Sum("estimate_hours"), act=Sum("actual_hours"))

    # カンバンはステータスごとの4列
    kanban = {
        "open": tasks.filter(status="open"),
        "in_progress": tasks.filter(status="in_progress"),
        "review": tasks.filter(status="review"),
        "done": tasks.filter(status="done"),
    }

    # 一覧タブの絞り込み
    status_filter = request.GET.get("status")
    filtered_tasks = tasks.filter(status=status_filter) if status_filter else tasks

    return render(request, "dev/project_detail.html", {
        "project": project,
        "tasks": filtered_tasks,
        "kanban": kanban,
        "total": total,
        "done": done,
        "in_progress": in_progress,
        "open_count": open_count,
        "review_count": review_count,
        "open_bugs": open_bugs,
        "estimate_total": hours["est"] or 0,
        "actual_total": hours["act"] or 0,
        "progress_pct": int(done / total * 100) if total > 0 else 0,
        "status_choices": DevTask.Status.choices,
    })


@owner_required
def project_create(request):
    if request.method == "POST":
        form = DevProjectForm(request.POST)
        if form.is_valid():
            project = form.save()
            messages.success(request, "プロジェクトを作成しました。")
            return redirect("manage:dev:project_detail", pk=project.pk)
    else:
        form = DevProjectForm()
    return render(request, "dev/project_form.html", {"form": form})


@owner_required
def project_edit(request, pk):
    project = get_object_or_404(DevProject, pk=pk)
    if request.method == "POST":
        form = DevProjectForm(request.POST, instance=project)
        if form.is_valid():
            form.save()
            messages.success(request, "プロジェクトを更新しました。")
            return redirect("manage:dev:project_detail", pk=project.pk)
    else:
        form = DevProjectForm(instance=project)
    return render(request, "dev/project_form.html", {"form": form})


@owner_required
def project_delete(request, pk):
    project = get_object_or_404(DevProject, pk=pk)
    if request.method == "POST":
        project.delete()
        messages.success(request, "プロジェクトを削除しました。")
        return redirect("manage:dev:project_list")
    return render(request, "dev/project_confirm_delete.html", {"project": project})


@owner_required
def task_create(request, project_pk):
    project = get_object_or_404(DevProject, pk=project_pk)
    if request.method == "POST":
        form = DevTaskForm(request.POST)
        if form.is_valid():
            task = form.save(commit=False)
            task.project = project
            # 表示順は末尾に付ける（KEM と同じく、いまの件数を順にする）
            task.sort_order = project.tasks.count()
            task.save()
            if task.assignee:
                notify_task_assigned(task, request.user)
            messages.success(request, "タスクを作成しました。")
            return redirect("manage:dev:project_detail", pk=project.pk)
    else:
        form = DevTaskForm()
    return render(request, "dev/task_form.html", {"form": form, "project": project})


@owner_required
def task_edit(request, pk):
    task = get_object_or_404(DevTask.objects.select_related("project"), pk=pk)
    old_assignee_id = task.assignee_id
    old_status = task.status
    if request.method == "POST":
        form = DevTaskForm(request.POST, instance=task)
        if form.is_valid():
            task = form.save()
            if task.assignee_id and task.assignee_id != old_assignee_id:
                notify_task_assigned(task, request.user)
            if task.status != old_status:
                notify_task_status_changed(task, request.user, old_status)
            messages.success(request, "タスクを更新しました。")
            return redirect("manage:dev:project_detail", pk=task.project.pk)
    else:
        form = DevTaskForm(instance=task)
    return render(request, "dev/task_form.html", {"form": form, "project": task.project})


@owner_required
def task_detail(request, pk):
    task = get_object_or_404(DevTask.objects.select_related("project", "assignee"), pk=pk)
    comments = task.comments.select_related("author").all()

    if request.method == "POST":
        comment_form = DevCommentForm(request.POST)
        if comment_form.is_valid():
            comment = comment_form.save(commit=False)
            comment.task = task
            comment.author = request.user
            comment.save()
            messages.success(request, "コメントを投稿しました。")
            return redirect("manage:dev:task_detail", pk=task.pk)
    else:
        comment_form = DevCommentForm()

    return render(request, "dev/task_detail.html", {
        "task": task,
        "project": task.project,
        "comments": comments,
        "comment_form": comment_form,
    })


@owner_required
def task_delete(request, pk):
    task = get_object_or_404(DevTask.objects.select_related("project"), pk=pk)
    project_pk = task.project.pk
    if request.method == "POST":
        task.delete()
        messages.success(request, "タスクを削除しました。")
        return redirect("manage:dev:project_detail", pk=project_pk)
    return render(request, "dev/task_confirm_delete.html", {"task": task})


@owner_required
def task_move(request, pk):
    """タスクのステータスを変更する（カンバンからの移動用）。"""
    task = get_object_or_404(DevTask.objects.select_related("project"), pk=pk)
    new_status = request.POST.get("status")
    if new_status and new_status in dict(DevTask.Status.choices):
        old_status = task.status
        task.status = new_status
        task.save(update_fields=["status", "updated_at"])
        notify_task_status_changed(task, request.user, old_status)
    # タスク一覧から進めたときは、絞り込んだままのその住所へ戻す（次の行にそのまま進める）
    戻り先 = request.POST.get("next") or ""
    if 戻り先 and url_has_allowed_host_and_scheme(
            戻り先, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
        return redirect(戻り先)
    return redirect("manage:dev:project_detail", pk=task.project.pk)


# ── 目安箱 ──────────────────────────────────────────────


@owner_required
def meyasubako_list(request):
    posts = Meyasubako.objects.select_related("reporter").all()
    return render(request, "dev/meyasubako_list.html", {"posts": posts})


@owner_required
def meyasubako_create(request):
    if request.method == "POST":
        form = MeyasubakoForm(request.POST, request.FILES)
        if form.is_valid():
            post = form.save(commit=False)
            post.reporter_name = post.reporter.get_full_name() or post.reporter.username
            shot = form.cleaned_data.get("screenshot")
            if shot:
                # 他の画像と同じく縮めて media/dev/ に置き、URL だけを持つ
                post.screenshot_url = images.保存する(shot, "dev")
            post.save()
            messages.success(request, "ご意見を投稿しました。ありがとうございます！")
            return redirect("manage:dev:meyasubako_list")
    else:
        # 入っている人を最初から選んでおく（KEM の「ログインユーザーの Worker を初期選択」と同じ）
        form = MeyasubakoForm(initial={"reporter": request.user.pk})
    return render(request, "dev/meyasubako_form.html", {"form": form})


@owner_required
def meyasubako_detail(request, pk):
    post = get_object_or_404(Meyasubako, pk=pk)
    return render(request, "dev/meyasubako_detail.html", {"post": post})


@owner_required
def meyasubako_resolve(request, pk):
    post = get_object_or_404(Meyasubako, pk=pk)
    if request.method == "POST":
        post.resolved = not post.resolved
        post.save(update_fields=["resolved", "updated_at"])
    return redirect("manage:dev:meyasubako_detail", pk=pk)
