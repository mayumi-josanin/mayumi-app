from django.contrib import admin

from .models import DevComment, DevProject, DevTask, Meyasubako


@admin.register(DevProject)
class DevProjectAdmin(admin.ModelAdmin):
    list_display = ("name", "status", "assignee", "start_date", "due_date")
    list_filter = ("status",)
    search_fields = ("name",)


@admin.register(DevTask)
class DevTaskAdmin(admin.ModelAdmin):
    list_display = ("title", "project", "status", "priority", "category", "assignee")
    list_filter = ("status", "priority", "category")
    search_fields = ("title",)


@admin.register(DevComment)
class DevCommentAdmin(admin.ModelAdmin):
    list_display = ("task", "author", "created_at")


@admin.register(Meyasubako)
class MeyasubakoAdmin(admin.ModelAdmin):
    list_display = ("title", "kind", "module", "reporter_name", "urgency", "resolved", "created_at")
    list_filter = ("kind", "module", "urgency", "resolved")
    search_fields = ("title", "reporter_name")
