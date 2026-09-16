"""開発管理の画面（KEM の apps/devkanri/views.py の写し）。

全部 **まゆみだけ**（owner_required）。スタッフには左メニューにも出るが 403。
KEM は login_required ＋ 会社（tenant）で絞っていた。こちらは会社が無いので絞りは要らない。
"""

from django.contrib import messages
from django.db.models import Count, Q, Sum
from django.shortcuts import get_object_or_404, redirect, render

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
