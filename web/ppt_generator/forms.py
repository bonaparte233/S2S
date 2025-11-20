"""
Forms for PPT Generator application.
"""

from django import forms
from .models import PPTGeneration


class PPTGenerationForm(forms.ModelForm):
    """Form for creating a new PPT generation request."""

    # Template selection options
    TEMPLATE_CHOICES = [
        ("default", "默认模板 (template.pptx)"),
        ("upload", "上传自定义模板"),
    ]

    # Config template selection options
    CONFIG_TEMPLATE_CHOICES = [
        ("auto", "自动匹配（根据 PPTX 模板）"),
        ("select", "从预设模板中选择"),
        ("upload", "上传自定义配置"),
    ]

    # LLM Provider options
    LLM_PROVIDER_CHOICES = [
        ("", "不使用大模型"),
        ("deepseek", "DeepSeek"),
        ("local", "本地部署模型"),
        ("custom", "自定义服务"),
    ]

    # DeepSeek model options
    DEEPSEEK_MODEL_CHOICES = [
        ("deepseek-chat", "DeepSeek Chat"),
        ("deepseek-reasoner", "DeepSeek Reasoner"),
    ]

    template_choice = forms.ChoiceField(
        choices=TEMPLATE_CHOICES,
        initial="default",
        widget=forms.RadioSelect,
        label="选择PPT模板",
        required=True,
    )

    config_template_choice = forms.ChoiceField(
        choices=CONFIG_TEMPLATE_CHOICES,
        initial="auto",
        widget=forms.RadioSelect,
        label="大模型配置模板 (template.json)",
        required=False,
    )

    llm_provider = forms.ChoiceField(
        choices=LLM_PROVIDER_CHOICES,
        initial="",
        widget=forms.Select(attrs={"class": "select-input"}),
        label="LLM供应商",
        required=False,
    )

    llm_model = forms.CharField(
        max_length=100,
        required=False,
        widget=forms.TextInput(
            attrs={
                "class": "text-input",
                "placeholder": "例如：deepseek-chat",
            }
        ),
        label="模型名称",
    )

    llm_api_key = forms.CharField(
        max_length=500,
        required=False,
        widget=forms.TextInput(
            attrs={
                "class": "text-input",
                "placeholder": "输入API Key",
                "type": "password",
            }
        ),
        label="API Key",
    )

    llm_base_url = forms.CharField(
        max_length=500,
        required=False,
        widget=forms.TextInput(
            attrs={
                "class": "text-input",
                "placeholder": "例如：https://api.deepseek.com 或 http://localhost:8000",
            }
        ),
        label="服务器地址",
    )

    user_prompt = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={
                "class": "textarea-input",
                "placeholder": "在此输入额外的提示词，将附加到系统默认提示词之后...",
                "rows": 3,
            }
        ),
        label="自定义Prompt（可选）",
    )

    class Meta:
        model = PPTGeneration
        fields = [
            "docx_file",
            "template_file",
            "config_template",
            "config_template_file",
            "use_llm",
            "llm_provider",
            "llm_model",
            "llm_api_key",
            "llm_base_url",
            "user_prompt",
            "course_name",
            "college_name",
            "lecturer_name",
        ]
        widgets = {
            "docx_file": forms.FileInput(
                attrs={
                    "class": "file-input",
                    "accept": ".docx",
                }
            ),
            "template_file": forms.FileInput(
                attrs={
                    "class": "file-input",
                    "accept": ".pptx",
                }
            ),
            "config_template_file": forms.FileInput(
                attrs={
                    "class": "file-input",
                    "accept": ".json",
                }
            ),
            "use_llm": forms.CheckboxInput(
                attrs={
                    "class": "checkbox-input",
                }
            ),
            "course_name": forms.TextInput(
                attrs={
                    "class": "text-input",
                    "placeholder": "例如：传感器技术基础",
                }
            ),
            "college_name": forms.TextInput(
                attrs={
                    "class": "text-input",
                    "placeholder": "例如：工程学院",
                }
            ),
            "lecturer_name": forms.TextInput(
                attrs={
                    "class": "text-input",
                    "placeholder": "例如：张教授",
                }
            ),
        }
        labels = {
            "docx_file": "讲稿文件 (DOCX)",
            "template_file": "自定义模板 (PPTX)",
            "config_template_file": "自定义配置模板 (JSON)",
            "use_llm": "启用大模型智能规划",
            "course_name": "课程名称（可选）",
            "college_name": "学院名称（可选）",
            "lecturer_name": "讲师名称（可选）",
        }
        help_texts = {
            "docx_file": "请上传包含讲稿内容的 Word 文档",
            "template_file": '仅在选择"上传自定义模板"时需要',
            "config_template_file": "上传自定义 JSON 配置模板（可选，优先级高于下拉选择）",
            "use_llm": "使用大模型自动规划幻灯片内容布局",
        }

    def clean(self):
        cleaned_data = super().clean()
        template_choice = cleaned_data.get("template_choice")
        template_file = cleaned_data.get("template_file")
        use_llm = cleaned_data.get("use_llm")
        llm_provider = cleaned_data.get("llm_provider")

        # Validate template file is provided when upload is selected
        if template_choice == "upload" and not template_file:
            raise forms.ValidationError('选择"上传自定义模板"时必须上传模板文件')

        # For developers: if use_llm is checked but no provider selected, set provider to empty
        # For non-developers: use_llm checkbox controls the behavior directly
        # Note: use_llm value comes from the checkbox, don't override it

        return cleaned_data
