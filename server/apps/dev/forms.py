"""開発管理のフォーム。KEM の apps/devkanri/forms.py と同じ項目。

KEM は担当者の候補を「社員番号 G の作業員」から作っていた。こちらに作業員は無く、
管理画面に入れるのは「まゆみ」と「スタッフ」だけなので、その2人（役割の group に入っている
ユーザー）を候補にする。
"""

from django import forms
from django.contrib.auth import get_user_model
from django.db.models import Q

from apps.manage.permissions import ROLES

from .models import DevComment, DevProject, DevTask, Meyasubako


def 管理画面の人たち():
    """担当者・責任者・報告者の候補。役割を持つ人と superuser（まゆみが superuser の箱もある）。"""
    User = get_user_model()
    return User.objects.filter(Q(groups__name__in=ROLES) | Q(is_superuser=True)).distinct().order_by("username")


def _form_control(form):
    """KEM と同じく、Textarea・日付・ファイル以外の部品に form-control を付ける。"""
    for field in form.fields.values():
        if not isinstance(field.widget, (forms.Textarea, forms.DateInput, forms.FileInput)):
            field.widget.attrs.setdefault("class", "form-control")


class DevProjectForm(forms.ModelForm):
    class Meta:
        model = DevProject
        fields = ["name", "description", "status", "assignee", "start_date", "due_date", "discord_webhook_url"]
        widgets = {
            "start_date": forms.DateInput(attrs={"type": "date", "class": "form-control"}),
            "due_date": forms.DateInput(attrs={"type": "date", "class": "form-control"}),
            "description": forms.Textarea(attrs={"class": "form-control", "rows": 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        _form_control(self)
        self.fields["assignee"].queryset = 管理画面の人たち()


class DevTaskForm(forms.ModelForm):
    class Meta:
        model = DevTask
        fields = [
            "title", "description", "status", "priority", "category",
            "assignee", "due_date", "estimate_hours", "actual_hours",
            "github_issue_url", "github_pr_url",
        ]
        widgets = {
            "due_date": forms.DateInput(attrs={"type": "date", "class": "form-control"}),
            "description": forms.Textarea(attrs={"class": "form-control", "rows": 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        _form_control(self)
        self.fields["assignee"].queryset = 管理画面の人たち()


class DevCommentForm(forms.ModelForm):
    class Meta:
        model = DevComment
        fields = ["body"]
        widgets = {
            "body": forms.Textarea(attrs={"class": "form-control", "rows": 2, "placeholder": "コメントを入力..."}),
        }


class MeyasubakoForm(forms.ModelForm):
    # KEM は作業員（Worker）から選んでいた。こちらは管理画面のユーザーから選ぶ
    reporter = forms.ModelChoiceField(queryset=None, label="報告者", empty_label="-- 選択してください --")
    # 画像はモデルに持たず、views で apps/manage/images.py に渡して URL にする
    screenshot = forms.ImageField(label="スクリーンショット", required=False)

    class Meta:
        model = Meyasubako
        fields = ["reporter", "kind", "module", "title", "problem", "wish", "urgency", "screenshot"]
        widgets = {
            "problem": forms.Textarea(attrs={
                "class": "form-control", "rows": 4, "placeholder": "何が起きましたか？ 何に困っていますか？",
            }),
            "wish": forms.Textarea(attrs={
                "class": "form-control", "rows": 3, "placeholder": "どうなると嬉しいですか？（任意）",
            }),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["reporter"].queryset = 管理画面の人たち()
        _form_control(self)
