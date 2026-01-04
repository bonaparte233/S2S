"""JSON 解析工具模块。"""

from __future__ import annotations

import json
from typing import Any, Dict, List


def _extract_json_value(text: str, opener: str) -> Any:
    """从文本中提取 JSON 值。"""
    decoder = json.JSONDecoder()
    idx = 0
    while idx < len(text):
        start = text.find(opener, idx)
        if start == -1:
            break
        try:
            value, _ = decoder.raw_decode(text[start:])
            return value
        except json.JSONDecodeError:
            idx = start + 1
    raise ValueError("模型输出中未找到 JSON")


def ensure_json_object(text: str) -> Dict:
    """确保文本解析为 JSON 对象。"""
    value = _extract_json_value(text.strip(), "{")
    if not isinstance(value, dict):
        raise ValueError("解析结果不是 JSON 对象")
    return value


def ensure_json_array(text: str) -> List[Dict]:
    """确保文本解析为 JSON 数组。"""
    value = _extract_json_value(text.strip(), "[")
    if not isinstance(value, list):
        raise ValueError("解析结果不是 JSON 数组")
    return value


def coerce_dict(entry: Any) -> Dict:
    """将各种类型强制转换为字典。"""
    if isinstance(entry, dict):
        return entry
    if isinstance(entry, str):
        entry = entry.strip()
        if not entry:
            raise ValueError("模型输出的元素为空字符串，无法解析为 JSON 对象。")
        if entry.startswith("{"):
            try:
                return json.loads(entry)
            except json.JSONDecodeError as exc:
                raise ValueError("模型输出的字符串不是合法 JSON 对象。") from exc
        raise ValueError("字符串元素必须是 JSON 对象字面量。")
    if isinstance(entry, list):
        for candidate in entry:
            try:
                return coerce_dict(candidate)
            except ValueError:
                continue
        raise ValueError("列表元素中未找到 JSON 对象。")
    raise ValueError("模型输出的元素不是有效的 JSON 对象。")
