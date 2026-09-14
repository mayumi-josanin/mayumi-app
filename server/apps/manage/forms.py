from django import forms

from apps.content.models import Category, News


class MultiFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class MultiFileField(forms.FileField):
    """複数の画像を一度に受ける。Django の FileField は1つしか受けないため。"""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("widget", MultiFileInput(attrs={"accept": "image/*", "multiple": True}))
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        single = super().clean
        if isinstance(data, (list, tuple)):
            return [single(d, initial) for d in data]
        return [single(data, initial)] if data else []


class NewsForm(forms.ModelForm):
    """お知らせの投稿・修正。GAS の管理アプリ（NEWS 配信・管理）と同じ項目。"""

    images = MultiFileField(
        label="画像を追加",
        required=False,
        help_text="スマートフォンのライブラリから選べます。大きな写真は自動で縮めます。",
    )

    send_push = forms.BooleanField(
        label="お客様のアプリに通知を送る（公開するときだけ）",
        required=False,
        help_text="「NEWSが更新されました」の通知が、アプリを入れている方に届きます。下書きや公開開始が先の日時のときは送りません。",
    )

    class Meta:
        model = News
        fields = [
            "posted_on",
            "title",
            "category",
            "icon",
            "body",
            "publish_at",
            "link_url",
            "button_text",
            "notice_listed",
        ]
        labels = {
            "posted_on": "日付",
            "publish_at": "公開開始日時（空欄ならすぐ公開）",
            "link_url": "リンクURL（任意）",
            "button_text": "リンクボタンの文字（任意）",
            "notice_listed": "お客様アプリの「お知らせ一覧」に出す",
        }
        widgets = {
            "posted_on": forms.DateInput(attrs={"type": "date"}),
            "publish_at": forms.DateTimeInput(attrs={"type": "datetime-local"}),
            "body": forms.Textarea(attrs={"rows": 12, "placeholder": "お知らせの内容を入力..."}),
            "icon": forms.TextInput(attrs={"placeholder": "📢"}),
            "link_url": forms.URLInput(attrs={"placeholder": "https://"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # カテゴリは表（Category）から。**種別はカテゴリ側が持つ**（お知らせ／ブログ）。
        names = list(Category.objects.order_by("sheet_row").values_list("name", "kind"))
        choices = [(n, f"{n}（{k}）" if k else n) for n, k in names]
        current = (self.instance.category or "").strip() if self.instance and self.instance.pk else ""
        if current and current not in [n for n, _ in names]:
            choices.append((current, current))
        self.fields["category"] = forms.ChoiceField(label="カテゴリ", choices=choices)
        self.fields["title"].widget.attrs["placeholder"] = "例：夏の産後ヨガ体験会"
        # 入力欄の見た目は KEM と同じ .form-control（チェックボックスは除く）
        for name, field in self.fields.items():
            if getattr(field.widget, "input_type", "") != "checkbox":
                field.widget.attrs["class"] = (field.widget.attrs.get("class", "") + " form-control").strip()
        self.fields["publish_at"].input_formats = ["%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S"]
