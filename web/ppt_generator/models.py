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
    llm_provider = models.CharField(
        max_length=50,
        choices=[
            ("deepseek", "DeepSeek"),
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
        help_text="例如：deepseek-chat, deepseek-reasoner, Qwen3-8B",
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
        return super().save(*args, **kwargs)

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
