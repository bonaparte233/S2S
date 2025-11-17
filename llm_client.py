"""大模型抽象、DeepSeek Provider 以及可扩展的本地/自定义 Provider 封装。"""

from __future__ import annotations

import os
import requests
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional


class BaseLLM(ABC):
    """大模型抽象基类，子类需实现 generate 方法。"""

    @abstractmethod
    def generate(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """传入类似 OpenAI Chat 的 messages，返回模型生成的文本。"""
        raise NotImplementedError


class OpenAILikeLLM(BaseLLM):
    """通用的 OpenAI Chat Completions 兼容接口。"""

    def __init__(
        self,
        model: str,
        base_url: str,
        api_key: Optional[str] = None,
        timeout: int = 60,
        extra_headers: Optional[Dict[str, str]] = None,
        completion_path: str = "/chat/completions",
    ):
        self.model = model
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.timeout = timeout
        self.extra_headers = extra_headers or {}
        self.completion_path = completion_path

    def generate(self, messages: List[Dict[str, str]], **kwargs) -> str:
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": messages,
        }
        if "temperature" in kwargs:
            payload["temperature"] = kwargs["temperature"]
        if "response_format" in kwargs:
            payload["response_format"] = kwargs["response_format"]

        headers = {"Content-Type": "application/json", **self.extra_headers}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        resp = requests.post(
            f"{self.base_url}{self.completion_path}",
            json=payload,
            headers=headers,
            timeout=self.timeout,
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"LLM API 调用失败：{resp.status_code} {resp.text}"
            )

        data = resp.json()
        try:
            return data["choices"][0]["message"]["content"]
        except (KeyError, IndexError) as exc:
            raise RuntimeError(f"LLM 返回格式异常：{data}") from exc


class DeepSeekLLM(OpenAILikeLLM):
    """DeepSeek Chat API 封装（OpenAI 风格接口）。"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = "deepseek-chat",
        base_url: Optional[str] = None,
        timeout: int = 60,
    ):
        key = api_key or os.getenv("DEEPSEEK_API_KEY")
        if not key:
            raise ValueError("未检测到 DEEPSEEK_API_KEY，请在环境变量中配置。")
        base = (
            base_url
            or os.getenv("DEEPSEEK_BASE_URL")
            or "https://api.deepseek.com"
        )
        super().__init__(model=model, base_url=base, api_key=key, timeout=timeout)


class LocalLLM(OpenAILikeLLM):
    """本地部署模型接口，兼容 OpenAI Chat Completions。"""

    def __init__(
        self,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        timeout: int = 60,
        api_key: Optional[str] = None,
    ):
        base = base_url or os.getenv("LOCAL_LLM_BASE_URL") or "http://127.0.0.1:8000/v1"
        model_name = model or os.getenv("LOCAL_LLM_MODEL") or "local-model"
        # 有些本地服务同样需要 key，可通过环境变量 LOCAL_LLM_API_KEY 指定
        key = api_key or os.getenv("LOCAL_LLM_API_KEY")
        super().__init__(model=model_name, base_url=base, api_key=key, timeout=timeout)
