"""
DeepSeek LLM Provider 实现

功能: BaseLLMProvider 的 DeepSeek 具体实现
职责: 只负责与 DeepSeek API 交互，不涉及其他业务逻辑 (SRP 原则)
"""

import json
import os
from openai import OpenAI
from .base_provider import BaseLLMProvider


class DeepSeekProvider(BaseLLMProvider):
    """
    DeepSeek API 的 LLM Provider 实现

    DeepSeek API 兼容 OpenAI 接口，使用 openai 库进行调用
    """

    def __init__(self, config):
        """
        初始化 DeepSeek Provider

        :param config: 配置字典，必须包含 'api_key' 或从环境变量读取
        """
        super().__init__(config)

        # 获取 API 密钥 (配置驱动原则)
        api_key = config.get("api_key") or os.environ.get("DEEPSEEK_API_KEY")
        if not api_key:
            raise ValueError(
                "DEEPSEEK_API_KEY 未设置。请在配置中提供 'api_key' 或设置环境变量 DEEPSEEK_API_KEY"
            )

        # 初始化 DeepSeek 客户端 (使用 OpenAI 兼容接口)
        self.client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        # 设置模型名称
        self.model_name = config.get("model_name", "deepseek-chat")

    def _build_prompt(self, raw_text: str, layout_placeholders: dict) -> str:
        """
        构建发送给 DeepSeek 的提示词 (私有方法)

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
        使用 DeepSeek API 处理幻灯片内容

        :param raw_text: 原始讲稿文本
        :param layout_placeholders: 占位符约束信息
        :return: 处理后的结构化内容字典
        :raises Exception: 当 API 调用失败时抛出异常
        """
        # 构建提示词
        prompt = self._build_prompt(raw_text, layout_placeholders)

        try:
            # 调用 DeepSeek API (使用 OpenAI 兼容接口)
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {
                        "role": "system",
                        "content": "你是一个专业的PPT内容撰写专家，擅长将口语化的讲稿转换为简洁、专业的PPT内容。",
                    },
                    {"role": "user", "content": prompt},
                ],
                response_format={"type": "json_object"},
                temperature=0.7,
                max_tokens=2000,
            )

            # 提取响应内容
            content = response.choices[0].message.content

            # 解析 JSON 响应
            return json.loads(content)

        except Exception as e:
            # 防御性编程: 向上抛出异常，由调度器捕获
            print(f"Error processing slide with DeepSeek: {e}")
            raise
