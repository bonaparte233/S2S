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
    path("developer-tools/", views.developer_tools, name="developer_tools"),
    path(
        "developer-tools/generate/",
        views.generate_config_template,
        name="generate_config_template",
    ),
    path(
        "developer-tools/ai-enrich/",
        views.ai_enrich_template_view,
        name="ai_enrich_template",
    ),
    # Template Editor - Independent Page
    path(
        "developer-tools/template-editor/",
        views.template_editor_page,
        name="template_editor_page",
    ),
    # Template Editor - API Endpoints
    path(
        "developer-tools/parse-ppt/",
        views.parse_ppt_template,
        name="parse_ppt_template",
    ),
    path(
        "developer-tools/update-shape-name/",
        views.update_shape_name_api,
        name="update_shape_name",
    ),
    path(
        "developer-tools/generate-config/",
        views.generate_template_config,
        name="generate_template_config",
    ),
    path(
        "developer-tools/download-ppt/<str:template_id>/",
        views.download_template_ppt,
        name="download_template_ppt",
    ),
    path(
        "developer-tools/toggle-shape-visibility/",
        views.toggle_shape_visibility,
        name="toggle_shape_visibility",
    ),
]
