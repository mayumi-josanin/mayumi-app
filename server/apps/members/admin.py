from django.contrib import admin

from .models import Member


@admin.register(Member)
class MemberAdmin(admin.ModelAdmin):
    """受付で会員情報を見て直すための画面。

    パスコードのハッシュは出さない。見ても意味が無いうえ、
    画面に出ていると「これを書き換えればよい」と誤解されるため。
    """

    list_display = ("member_id", "name", "kana", "phone", "stamp_count", "bijiris_registered", "deleted_at")
    list_filter = ("bijiris_registered", "push_enabled", "role")
    search_fields = ("member_id", "name", "kana", "phone")
    ordering = ("member_id",)
    readonly_fields = ("member_id", "imported_at", "changed_at")
    exclude = ("passcode_hash", "password_hash", "password_salt")
