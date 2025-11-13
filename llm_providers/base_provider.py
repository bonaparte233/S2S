"""
LLM Provider 抽象基类

功能: 定义 LLM 插件的抽象接口，实现依赖倒置原则 (DIP)
设计原则: 高层模块 (process_with_llm.py) 只依赖此抽象，不依赖具体实现
"""

from abc import ABC, abstractmethod


class BaseLLMProvider(ABC):
    """
    LLM Provider 抽象基类
    
    所有 LLM 实现 (如 GeminiProvider, OpenAIProvider) 都必须继承此类
    并实现 process_slide 方法
    """
    
    def __init__(self, config=None):
        """
        初始化 LLM Provider
        
        :param config: 配置字典，包含 API 密钥、模型名称等信息
        """
        self.config = config or {}
    
    @abstractmethod
    def process_slide(self, raw_text: str, layout_placeholders: dict) -> dict:
        """
        处理单页幻灯片的原始文本，并根据约束返回结构化内容
        
        :param raw_text: S1 传入的原始文本 (包含 Markdown 和 [IMAGE: ...] 标记)
        :param layout_placeholders: map.json 中该布局的 'placeholders' 对象
                                   格式: {"title": {"id": 0, "constraints": {...}}, ...}
        :return: 字典，键为 'placeholders' 中的键名 (如 "title", "body_text"),
                值为 LLM 处理后的内容
                格式: {"title": "处理后的标题", "body_text": "处理后的正文", ...}
        """
        pass

