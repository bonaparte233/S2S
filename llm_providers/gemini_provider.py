"""
Google Gemini LLM Provider 实现

功能: BaseLLMProvider 的 Google Gemini 具体实现
职责: 只负责与 Gemini API 交互，不涉及其他业务逻辑 (SRP 原则)
"""

import google.generativeai as genai
import json
import os
from .base_provider import BaseLLMProvider


class GeminiProvider(BaseLLMProvider):
    """
    Google Gemini API 的 LLM Provider 实现
    """

    def __init__(self, config):
        """
        初始化 Gemini Provider

        :param config: 配置字典，必须包含 'api_key' 或从环境变量读取
        """
        super().__init__(config)

        # 获取 API 密钥 (配置驱动原则)
        api_key = config.get("api_key") or os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            raise ValueError(
                "GOOGLE_API_KEY 未设置。请在配置中提供 'api_key' 或设置环境变量 GOOGLE_API_KEY"
            )

        # 配置 Gemini API
        genai.configure(api_key=api_key)

        # 初始化模型
        model_name = config.get("model_name", "gemini-2.0-flash-exp")
        self.model = genai.GenerativeModel(model_name)

    def _build_prompt(self, raw_text: str, layout_placeholders: dict) -> str:
        """
        构建发送给 Gemini 的提示词 (私有方法)

        :param raw_text: 原始讲稿文本
        :param layout_placeholders: 占位符约束信息
        :return: 完整的提示词字符串
        """
        prompt_lines = [
            "你是一个专业的PPT内容撰写专家。你的任务是将口语化的讲稿转换为简洁、专业的PPT内容，并严格按照JSON格式返回。",
            "",
            "## 核心规则:",
            "1. **严格遵守字数和行数限制**: 每个字段都有 max_chars（最大字符数）和 max_lines（最大行数）限制，**绝对不能超过**。",
            "2. **保留 Markdown 列表格式**: 原文中的列表格式（如 '- ' 和 '  - '）必须保留。",
            "3. **表格转列表**: 如果原文包含 Markdown 表格，必须转换为多级列表，**严禁**输出表格。",
            "4. **图片处理**: ",
            "   - 如果有 `picture` 字段，从 `[IMAGE: path]` 标记中提取路径填入该字段",
            "   - 同时**必须删除**文本字段中的 `[IMAGE: ...]` 标记",
            "   - 如果有多个图片，只取第一个",
            "5. **内容精简**: 将口语化的讲稿转换为简洁、专业的PPT语言。",
            "6. **纯 JSON 输出**: 回复**只能**是 JSON 对象，不要任何解释文字或 markdown 标记（如 ```json）。",
            "",
            "## 字段约束详情:",
        ]

        # 收集所有输出键名并生成详细约束说明
        output_keys = []

        for key, details in layout_placeholders.items():
            output_keys.append(key)
            constraints = details.get("constraints", {})
            max_chars = constraints.get("max_chars", "未限制")
            max_lines = constraints.get("max_lines", "未限制")

            # 生成更清晰的约束说明
            prompt_lines.append(
                f"- **{key}**: 最多 {max_chars} 个字符，最多 {max_lines} 行"
            )

        # 添加原始文本和输出格式说明
        prompt_lines.extend(
            [
                "",
                "## 原始讲稿文本:",
                raw_text,
                "",
                "## 输出要求:",
                f"返回一个 JSON 对象，包含以下键: {output_keys}",
                "每个键的值必须是字符串类型。",
                "",
                "## 你的输出 (纯 JSON):",
            ]
        )

        return "\n".join(prompt_lines)

    def process_slide(self, raw_text: str, layout_placeholders: dict) -> dict:
        """
        使用 Gemini API 处理幻灯片内容

        :param raw_text: 原始讲稿文本
        :param layout_placeholders: 占位符约束信息
        :return: 处理后的结构化内容字典
        :raises Exception: 当 API 调用失败时抛出异常
        """
        # 构建提示词
        prompt = self._build_prompt(raw_text, layout_placeholders)

        # 配置生成参数 (要求返回 JSON 格式)
        generation_config = genai.GenerationConfig(
            response_mime_type="application/json"
        )

        try:
            # 调用 Gemini API
            response = self.model.generate_content(
                prompt, generation_config=generation_config
            )

            # 解析 JSON 响应
            return json.loads(response.text)

        except Exception as e:
            # 防御性编程: 向上抛出异常，由调度器捕获
            print(f"Error processing slide with Gemini: {e}")
            raise
