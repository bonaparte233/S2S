"""
URL configuration for PPT Generator app.
"""

from django.urls import path
from . import views

urlpatterns = [
    # Authentication
    path("login/", views.user_login, name="login"),
    path("logout/", views.user_logout, name="logout"),
    # Main pages
    path("", views.index, name="index"),
    path("generation/<int:pk>/", views.generation_detail, name="generation_detail"),
    path("generation/<int:pk>/start/", views.start_generation, name="start_generation"),
    path("generation/<int:pk>/status/", views.check_status, name="check_status"),
    path("generation/<int:pk>/download/", views.download_ppt, name="download_ppt"),
    path("history/", views.history, name="history"),
    # Developer only
    path("export-template/", views.export_template_json, name="export_template"),
]
