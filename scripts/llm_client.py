"""大模型抽象、DeepSeek Provider 以及可扩展的本地/自定义 Provider 封装。"""

from __future__ import annotations

import os
import requests
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional


class BaseLLM(ABC):
    """大模型抽象基类，子类需实现 generate 方法。"""

    @abstractmethod
    def generate(self, messages: List[Dict[str, Any]], **kwargs) -> str:
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

    def generate(self, messages: List[Dict[str, Any]], **kwargs) -> str:
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
        base = base_url or os.getenv("LOCAL_LLM_BASE_URL") or "http://172.18.75.58:9000/generate"
        model_name = model or os.getenv("LOCAL_LLM_MODEL") or "Qwen3-8B"
        # 有些本地服务同样需要 key，可通过环境变量 LOCAL_LLM_API_KEY 指定
        key = api_key or os.getenv("LOCAL_LLM_API_KEY")
        super().__init__(model=model_name, base_url=base, api_key=key, timeout=timeout)


class TaichuLLM(OpenAILikeLLM):
    """太初多模态大模型 API 封装。"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = "taichu_vl",
        base_url: Optional[str] = None,
        timeout: int = 60,
    ):
        key = api_key or os.getenv("TAICHU_API_KEY")
        if not key:
            raise ValueError("未检测到 TAICHU_API_KEY，请在环境变量中配置。")
        base = (
            base_url
            or os.getenv("TAICHU_BASE_URL")
            or "https://platform.wair.ac.cn/maas/v1"
        )
        model_name = model
        super().__init__(model=model_name, base_url=base, api_key=key, timeout=timeout)


class QwenVLLM(BaseLLM):
    """适配 vLLM 部署的 Qwen /generate 接口，自动拼接 chat 模板。"""

    def __init__(
        self,
        base_url: str,
        timeout: int = 99999,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

    @staticmethod
    def _format_messages(messages: List[Dict[str, str]]) -> str:
        def map_role(role: str) -> str:
            if role == "system":
                return "system"
            if role == "assistant":
                return "assistant"
            return "user"

        parts: List[str] = []
        for msg in messages:
            role = map_role(msg.get("role", "user"))
            content = msg.get("content", "")
            parts.append(f"<|im_start|>{role}\n{content}<|im_end|>\n")
        parts.append("<|im_start|>assistant\n")
        return "".join(parts)

    def generate(self, messages: List[Dict[str, Any]], **kwargs) -> str:
        prompt = self._format_messages(messages)
        payload = {
            "prompt": prompt,
            "max_tokens": kwargs.get("max_tokens", 5000),
            "temperature": kwargs.get("temperature", 0.3),
        }
        response = requests.post(
            f"{self.base_url}/generate",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=self.timeout,
        )
        if response.status_code != 200:
            raise RuntimeError(f"Qwen vLLM 调用失败：{response.status_code} {response.text}")
        data = response.json()
        print("⚠️ Qwen返回：", data)
        texts = data.get("text")
        if isinstance(texts, list) and texts:
            return texts[0]
        if isinstance(texts, str):
            return texts
        raise RuntimeError(f"Qwen vLLM 返回格式异常：{data}")
