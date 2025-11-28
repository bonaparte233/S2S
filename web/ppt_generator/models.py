"""
Models for PPT Generator application.
"""

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone


class GlobalLLMConfig(models.Model):
    """全局LLM配置 - 支持多个配置，可选择默认配置"""

    name = models.CharField(
        max_length=100,
        unique=True,
        verbose_name="配置名称",
        help_text="为此配置起一个易于识别的名称",
    )
    is_default = models.BooleanField(
        default=False,
        verbose_name="默认配置",
        help_text="勾选后，此配置将成为系统默认配置",
    )
    is_multimodal_default = models.BooleanField(
        default=False,
        verbose_name="默认多模态配置",
        help_text="勾选后，此配置将成为多模态（图像理解）任务的默认配置",
    )
    supports_multimodal = models.BooleanField(
        default=False,
        verbose_name="支持多模态",
        help_text="勾选表示此模型支持图像理解（如 GLM-4V、Taichu-VL 等视觉模型）",
    )
    llm_provider = models.CharField(
        max_length=50,
        choices=[
            ("deepseek", "DeepSeek"),
            ("taichu", "紫东太初多模态模型"),
            ("glm", "智谱AI (GLM)"),
            ("local", "本地部署模型"),
            ("custom", "自定义服务"),
        ],
        default="deepseek",
        verbose_name="LLM供应商",
    )
    llm_model = models.CharField(
        max_length=100,
        default="deepseek-chat",
        verbose_name="LLM模型",
        help_text="DeepSeek: deepseek-chat | 紫东太初: taichu4_vl_32b | 智谱AI: glm-4.6 | 本地: 自定义模型名称",
    )
    llm_api_key = models.CharField(
        max_length=500, blank=True, verbose_name="API Key", help_text="默认API密钥"
    )
    llm_base_url = models.CharField(
        max_length=500,
        blank=True,
        verbose_name="服务器地址",
        help_text="自定义服务器地址（可选）",
    )
    default_prompt = models.TextField(
        blank=True,
        verbose_name="默认系统Prompt",
        help_text="全局默认的系统提示词（可选）",
    )

    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新时间")
    updated_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        verbose_name="更新者",
    )

    class Meta:
        verbose_name = "全局LLM配置"
        verbose_name_plural = "全局LLM配置"
        ordering = ["-is_default", "name"]

    def save(self, *args, **kwargs):
        # 如果设置为默认配置，取消其他配置的默认状态
        if self.is_default:
            GlobalLLMConfig.objects.filter(is_default=True).exclude(pk=self.pk).update(
                is_default=False
            )
        # 如果这是第一个配置，自动设为默认
        elif not self.pk and not GlobalLLMConfig.objects.exists():
            self.is_default = True

        # 多模态默认配置：必须勾选了"支持多模态"
        if self.is_multimodal_default:
            if not self.supports_multimodal:
                self.is_multimodal_default = False
            else:
                # 取消其他配置的多模态默认状态
                GlobalLLMConfig.objects.filter(is_multimodal_default=True).exclude(
                    pk=self.pk
                ).update(is_multimodal_default=False)

        return super().save(*args, **kwargs)

    def get_model_for_provider(self):
        """根据提供商返回正确的模型名称"""
        # 如果用户明确设置了模型名称，优先使用
        if self.llm_model:
            # 检查是否是默认的 deepseek-chat，如果是且提供商不是 deepseek，则使用提供商的默认值
            if self.llm_model == "deepseek-chat" and self.llm_provider != "deepseek":
                return self._get_default_model_for_provider()
            return self.llm_model

        # 否则返回提供商的默认模型
        return self._get_default_model_for_provider()

    def _get_default_model_for_provider(self):
        """返回提供商的默认模型名称"""
        defaults = {
            "deepseek": "deepseek-chat",
            "taichu": "taichu4_vl_32b",
            "glm": "glm-4.6",
            "local": "local-model",
            "custom": "custom-model",
        }
        return defaults.get(self.llm_provider, "deepseek-chat")

    @classmethod
    def get_config(cls):
        """获取默认配置（如果不存在则创建默认配置）"""
        # 尝试获取默认配置
        config = cls.objects.filter(is_default=True).first()
        if config:
            return config

        # 如果没有默认配置，尝试获取第一个配置
        config = cls.objects.first()
        if config:
            config.is_default = True
            config.save()
            return config

        # 如果没有任何配置，创建默认配置
        config = cls.objects.create(
            name="默认配置",
            is_default=True,
            llm_provider="deepseek",
            llm_model="deepseek-chat",
        )
        return config

    @classmethod
    def get_multimodal_config(cls):
        """获取默认多模态配置"""
        # 尝试获取多模态默认配置
        config = cls.objects.filter(is_multimodal_default=True).first()
        if config:
            return config

        # 如果没有多模态默认配置，尝试获取第一个支持多模态的配置
        config = cls.objects.filter(supports_multimodal=True).first()
        if config:
            return config

        # 没有可用的多模态配置
        return None

    def __str__(self):
        default_mark = " [默认]" if self.is_default else ""
        return f"{self.name} ({self.llm_provider} - {self.llm_model}){default_mark}"


class PPTGeneration(models.Model):
    """Track PPT generation history."""

    STATUS_CHOICES = [
        ("pending", "等待中"),
        ("processing", "处理中"),
        ("completed", "已完成"),
        ("failed", "失败"),
    ]

    # User who created this generation
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="ppt_generations",
        verbose_name="用户",
    )

    # Input files
    docx_file = models.FileField(upload_to="uploads/docx/", verbose_name="讲稿文件")
    template_file = models.FileField(
        upload_to="uploads/templates/", null=True, blank=True, verbose_name="模板文件"
    )
    template_name = models.CharField(
        max_length=255, null=True, blank=True, verbose_name="模板名称"
    )
    config_template = models.CharField(
        max_length=500, null=True, blank=True, verbose_name="配置模板"
    )
    config_template_file = models.FileField(
        upload_to="uploads/config_templates/",
        null=True,
        blank=True,
        verbose_name="配置模板文件",
    )

    # Configuration
    use_llm = models.BooleanField(default=False, verbose_name="使用大模型")

    # LLM 配置选择方式
    LLM_CONFIG_CHOICES = [
        ("preset", "使用预设配置"),
        ("custom", "自定义配置"),
    ]
    llm_config_choice = models.CharField(
        max_length=20,
        choices=LLM_CONFIG_CHOICES,
        default="preset",
        verbose_name="配置方式",
    )

    # 预设配置（从 GlobalLLMConfig 中选择）
    llm_preset_config = models.ForeignKey(
        GlobalLLMConfig,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        verbose_name="预设配置",
        help_text="选择管理员预设的 LLM 配置",
    )

    # 自定义配置字段（仅在选择"自定义配置"时使用）
    llm_provider = models.CharField(
        max_length=50, null=True, blank=True, verbose_name="LLM供应商"
    )
    llm_model = models.CharField(
        max_length=100, null=True, blank=True, verbose_name="LLM模型"
    )
    llm_api_key = models.CharField(
        max_length=500, null=True, blank=True, verbose_name="API Key"
    )
    llm_base_url = models.CharField(
        max_length=500, null=True, blank=True, verbose_name="服务器地址"
    )
    user_prompt = models.TextField(
        null=True, blank=True, verbose_name="用户自定义Prompt"
    )
    course_name = models.CharField(
        max_length=255, null=True, blank=True, verbose_name="课程名称"
    )
    college_name = models.CharField(
        max_length=255, null=True, blank=True, verbose_name="学院名称"
    )
    lecturer_name = models.CharField(
        max_length=255, null=True, blank=True, verbose_name="讲师名称"
    )

    # Output
    output_ppt = models.FileField(
        upload_to="outputs/", null=True, blank=True, verbose_name="生成的PPT"
    )
    config_json = models.FileField(
        upload_to="configs/", null=True, blank=True, verbose_name="配置JSON"
    )

    # Status tracking
    status = models.CharField(
        max_length=20, choices=STATUS_CHOICES, default="pending", verbose_name="状态"
    )
    error_message = models.TextField(null=True, blank=True, verbose_name="错误信息")

    # Timestamps
    created_at = models.DateTimeField(default=timezone.now, verbose_name="创建时间")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新时间")
    completed_at = models.DateTimeField(null=True, blank=True, verbose_name="完成时间")

    # Run directory path
    run_dir = models.CharField(
        max_length=500, null=True, blank=True, verbose_name="运行目录"
    )

    class Meta:
        permissions = [
            ("is_developer", "开发者权限"),
            ("can_export_template_json", "可以导出模板JSON"),
            ("can_view_llm_config", "可以查看LLM配置"),
        ]
        verbose_name = "PPT生成记录"
        verbose_name_plural = "PPT生成记录"
        ordering = ["-created_at"]

    def __str__(self):
        return f"PPT生成 #{self.id} - {self.get_status_display()} - {self.created_at.strftime('%Y-%m-%d %H:%M')}"

    def mark_processing(self):
        """Mark as processing."""
        self.status = "processing"
        self.save(update_fields=["status", "updated_at"])

    def mark_completed(self, output_path, config_path=None, run_dir=None):
        """Mark as completed with output files."""
        self.status = "completed"
        self.completed_at = timezone.now()
        self.output_ppt = output_path
        if config_path:
            self.config_json = config_path
        if run_dir:
            self.run_dir = str(run_dir)
        self.save(
            update_fields=[
                "status",
                "completed_at",
                "output_ppt",
                "config_json",
                "run_dir",
                "updated_at",
            ]
        )

    def mark_failed(self, error_msg):
        """Mark as failed with error message."""
        self.status = "failed"
        self.error_message = error_msg
        self.save(update_fields=["status", "error_message", "updated_at"])


class TemplateEditSession(models.Model):
    """模板编辑会话记录 - 用于保存和恢复编辑进度"""

    EDITOR_TYPE_CHOICES = [
        ("ppt", "PPT 模板编辑器"),
        ("config", "配置模板编辑器"),
        ("wizard", "模板制作向导"),
    ]

    # 用户
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="template_edit_sessions",
        verbose_name="用户",
    )

    # 编辑器类型
    editor_type = models.CharField(
        max_length=20,
        choices=EDITOR_TYPE_CHOICES,
        default="ppt",
        verbose_name="编辑器类型",
    )

    # 会话 ID（对应 media/template_editor/{session_id} 或前端保存的数据）
    session_id = models.CharField(
        max_length=100,
        verbose_name="会话 ID",
        help_text="PPT 编辑器使用 UUID，配置编辑器使用前端生成的 ID",
    )

    # 模板名称（从文件名提取）
    template_name = models.CharField(
        max_length=255,
        verbose_name="模板名称",
    )

    # 编辑进度（JSON 格式存储状态）
    progress_data = models.JSONField(
        default=dict,
        blank=True,
        verbose_name="编辑进度数据",
        help_text="存储编辑进度，如已命名元素数、总元素数等",
    )

    # 缩略图路径（可选，用于预览）
    thumbnail_url = models.CharField(
        max_length=500,
        blank=True,
        null=True,
        verbose_name="缩略图 URL",
    )

    # 时间戳
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="创建时间")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新时间")

    class Meta:
        verbose_name = "模板编辑会话"
        verbose_name_plural = "模板编辑会话"
        ordering = ["-updated_at"]
        # 同一用户同一会话 ID 只能有一条记录
        unique_together = ["user", "session_id"]

    def __str__(self):
        return f"{self.get_editor_type_display()} - {self.template_name} ({self.user.username})"

    @property
    def progress_summary(self):
        """返回进度摘要字符串"""
        data = self.progress_data or {}
        if self.editor_type == "ppt":
            named = data.get("named_count", 0)
            total = data.get("total_count", 0)
            if total > 0:
                percent = int(named / total * 100)
                return f"{named}/{total} 已命名 ({percent}%)"
            return "未开始"
        elif self.editor_type == "config":
            pages = data.get("page_count", 0)
            filled = data.get("filled_count", 0)
            return f"{pages} 页, {filled} 个字段已填充"
        return ""
