"""
Admin configuration for PPT Generator.
"""

from django.contrib import admin
from django.utils.html import format_html
from .models import GlobalLLMConfig, PPTGeneration


@admin.register(GlobalLLMConfig)
class GlobalLLMConfigAdmin(admin.ModelAdmin):
    """Admin interface for Global LLM Configuration."""

    list_display = [
        "name",
        "is_default_badge",
        "supports_multimodal_badge",
        "llm_provider",
        "llm_model",
        "has_api_key",
        "updated_at",
        "updated_by",
    ]
    list_filter = [
        "is_default",
        "is_multimodal_default",
        "supports_multimodal",
        "llm_provider",
    ]
    search_fields = ["name", "llm_model"]
    readonly_fields = ["updated_at", "updated_by"]
    actions = ["set_as_default", "set_as_multimodal_default"]

    fieldsets = [
        (
            "配置标识",
            {
                "fields": ["name", "is_default", "is_multimodal_default"],
                "description": "为配置命名，选择默认配置类型",
            },
        ),
        (
            "基本配置",
            {
                "fields": ["llm_provider", "llm_model", "supports_multimodal"],
                "description": "配置LLM供应商和模型，勾选「支持多模态」表示此模型支持图像理解",
            },
        ),
        (
            "认证信息",
            {
                "fields": ["llm_api_key", "llm_base_url"],
                "description": "配置API密钥和服务器地址（如需要）",
            },
        ),
        (
            "高级选项",
            {
                "fields": ["default_prompt"],
                "classes": ["collapse"],
                "description": "配置全局默认的系统提示词",
            },
        ),
        (
            "元信息",
            {
                "fields": ["updated_at", "updated_by"],
                "classes": ["collapse"],
            },
        ),
    ]

    def is_default_badge(self, obj):
        """显示默认配置状态（包括普通默认和多模态默认）"""
        badges = []
        if obj.is_default:
            badges.append(
                '<span style="background-color: #28a745; color: white; padding: 3px 8px; border-radius: 3px; font-weight: bold; margin-right: 4px;">✓ 默认</span>'
            )
        if obj.is_multimodal_default:
            badges.append(
                '<span style="background-color: #9c27b0; color: white; padding: 3px 8px; border-radius: 3px; font-weight: bold;">🖼 多模态默认</span>'
            )
        if badges:
            return format_html("".join(badges))
        return format_html('<span style="color: #999;">-</span>')

    is_default_badge.short_description = "默认配置"

    def supports_multimodal_badge(self, obj):
        """显示是否支持多模态"""
        if obj.supports_multimodal:
            return format_html(
                '<span style="color: #9c27b0; font-weight: bold;">✓ 支持</span>'
            )
        return format_html('<span style="color: #999;">-</span>')

    supports_multimodal_badge.short_description = "多模态"

    def has_api_key(self, obj):
        """显示是否配置了API密钥"""
        if obj.llm_api_key:
            return format_html('<span style="color: green;">✓ 已配置</span>')
        return format_html('<span style="color: orange;">✗ 未配置</span>')

    has_api_key.short_description = "API密钥"

    def set_as_default(self, request, queryset):
        """将选中的配置设为默认配置"""
        if queryset.count() != 1:
            self.message_user(request, "请只选择一个配置设为默认", level="error")
            return

        config = queryset.first()
        # 取消其他配置的默认状态
        GlobalLLMConfig.objects.filter(is_default=True).update(is_default=False)
        # 设置当前配置为默认
        config.is_default = True
        config.save()

        self.message_user(
            request, f"已将 '{config.name}' 设为默认配置", level="success"
        )

    set_as_default.short_description = "设为默认配置"

    def set_as_multimodal_default(self, request, queryset):
        """将选中的配置设为多模态默认配置"""
        if queryset.count() != 1:
            self.message_user(request, "请只选择一个配置设为多模态默认", level="error")
            return

        config = queryset.first()
        if not config.supports_multimodal:
            self.message_user(
                request,
                f"配置 '{config.name}' 未勾选「支持多模态」，请先勾选后再设为多模态默认",
                level="error",
            )
            return

        # 取消其他配置的多模态默认状态
        GlobalLLMConfig.objects.filter(is_multimodal_default=True).update(
            is_multimodal_default=False
        )
        # 设置当前配置为多模态默认
        config.is_multimodal_default = True
        config.save()

        self.message_user(
            request, f"已将 '{config.name}' 设为多模态默认配置", level="success"
        )

    set_as_multimodal_default.short_description = "设为多模态默认配置"

    def save_model(self, request, obj, form, change):
        """保存时记录更新者"""
        obj.updated_by = request.user
        super().save_model(request, obj, form, change)


@admin.register(PPTGeneration)
class PPTGenerationAdmin(admin.ModelAdmin):
    """Admin interface for PPT Generation records."""

    list_display = [
        "id",
        "user_link",
        "course_name_short",
        "status_badge",
        "llm_status",
        "created_at",
        "completed_at",
    ]
    list_filter = ["status", "use_llm", "llm_provider", "created_at", "user"]
    search_fields = ["course_name", "college_name", "lecturer_name", "user__username"]
    readonly_fields = [
        "created_at",
        "updated_at",
        "completed_at",
        "run_dir",
        "status_badge",
        "download_links",
    ]

    # 每页显示数量
    list_per_page = 20

    # 日期层级导航
    date_hierarchy = "created_at"

    # 默认排序
    ordering = ["-created_at"]

    fieldsets = [
        ("用户信息", {"fields": ["user"], "description": "创建此生成任务的用户"}),
        (
            "输入文件",
            {
                "fields": ["docx_file", "template_file", "template_name"],
                "classes": ["collapse"],
            },
        ),
        (
            "课程信息",
            {
                "fields": ["course_name", "college_name", "lecturer_name"],
                "classes": ["wide"],
            },
        ),
        (
            "大模型配置",
            {
                "fields": [
                    "use_llm",
                    "llm_provider",
                    "llm_model",
                    "llm_api_key",
                    "llm_base_url",
                    "user_prompt",
                ],
                "classes": ["collapse"],
                "description": "LLM相关配置（仅开发者可见）",
            },
        ),
        (
            "输出文件",
            {
                "fields": ["output_ppt", "config_json", "run_dir", "download_links"],
                "classes": ["wide"],
            },
        ),
        (
            "状态信息",
            {
                "fields": [
                    "status",
                    "status_badge",
                    "error_message",
                    "created_at",
                    "updated_at",
                    "completed_at",
                ],
                "classes": ["wide"],
            },
        ),
    ]

    def user_link(self, obj):
        """显示用户名（带链接）"""
        if obj.user:
            return format_html(
                '<a href="/admin/auth/user/{}/change/">{}</a>',
                obj.user.id,
                obj.user.username,
            )
        return "-"

    user_link.short_description = "用户"

    def course_name_short(self, obj):
        """显示课程名称（截断）"""
        if obj.course_name:
            return (
                obj.course_name[:30] + "..."
                if len(obj.course_name) > 30
                else obj.course_name
            )
        return "-"

    course_name_short.short_description = "课程名称"

    def status_badge(self, obj):
        """显示状态徽章"""
        colors = {
            "pending": "#FFA500",
            "processing": "#1E90FF",
            "completed": "#28A745",
            "failed": "#DC3545",
        }
        color = colors.get(obj.status, "#6C757D")
        return format_html(
            '<span style="background-color: {}; color: white; padding: 3px 10px; '
            'border-radius: 3px; font-weight: bold;">{}</span>',
            color,
            obj.get_status_display(),
        )

    status_badge.short_description = "状态"

    def llm_status(self, obj):
        """显示LLM使用状态"""
        if obj.use_llm:
            provider = obj.llm_provider or "未知"
            return format_html('<span style="color: #28A745;">✓ {}</span>', provider)
        return format_html('<span style="color: #6C757D;">✗ 未使用</span>')

    llm_status.short_description = "LLM"

    def download_links(self, obj):
        """显示下载链接"""
        links = []
        if obj.output_ppt:
            links.append(
                format_html(
                    '<a href="{}" target="_blank" style="margin-right: 10px;">'
                    "📄 下载PPT</a>",
                    obj.output_ppt.url,
                )
            )
        if obj.config_json:
            links.append(
                format_html(
                    '<a href="{}" target="_blank">📋 下载JSON</a>', obj.config_json.url
                )
            )
        return format_html(" ".join(links)) if links else "-"

    download_links.short_description = "下载"

    def get_queryset(self, request):
        """优化查询性能"""
        qs = super().get_queryset(request)
        return qs.select_related("user")

    def has_add_permission(self, request):
        """禁止在admin中直接添加记录"""
        return False
