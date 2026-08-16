from django.urls import path

from . import views

urlpatterns = [
    path("health", views.health, name="health"),
    # 道は英字にする。PowerShell の curl で打つときに日本語だと化けるため。
    path("entry-info", views.入口の見え方, name="entry-info"),
    path("measurements", views.measurements, name="measurements"),
]
