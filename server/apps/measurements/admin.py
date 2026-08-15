from django.contrib import admin

from .models import Measurement


@admin.register(Measurement)
class MeasurementAdmin(admin.ModelAdmin):
    list_display = ("customer_name", "measured_on", "waist", "hip", "thigh_right", "thigh_left", "member_id")
    list_filter = ("measured_on",)
    search_fields = ("customer_name", "measurement_id", "member_id")
    ordering = ("customer_name", "measured_on")
