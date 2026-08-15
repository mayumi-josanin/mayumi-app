from django.urls import path

from . import views

urlpatterns = [
    path("health", views.health, name="health"),
    path("measurements", views.measurements, name="measurements"),
]
