"""模板定义处理模块。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def _collect_fields(schema: Any, prefix: Optional[Tuple] = None) -> List[Dict]:
    """递归收集模板 schema 中的所有字段。"""
    results = []
    prefix = prefix or ()

    if isinstance(schema, dict):
        if "type" in schema and "value" in schema:
            field_type = schema.get("type") or "text"
            results.append({
                "path": prefix,
                "is_image": field_type.lower() == "image",
                "hint": schema.get("hint") or "",
                "max_chars": schema.get("max_chars"),
                "required": bool(schema.get("required", False)),
            })
        else:
            for key, value in schema.items():
                results.extend(_collect_fields(value, prefix + (key,)))
    else:
        results.append({
            "path": prefix,
            "is_image": any("图片" in seg or "image" in seg.lower() for seg in prefix),
            "hint": "",
            "max_chars": None,
            "required": False,
        })

    return results


def load_template_defs(
    template_json: str, template_list: Optional[str] = None
) -> Tuple[Dict[int, Dict], Dict]:
    """加载模板定义和全局配置。

    Args:
        template_json: template.json 文件路径
        template_list: （已废弃）template.txt 文件路径

    Returns:
        (templates, global_config):
        - templates: 模板定义字典，key 为模板页码
        - global_config: 全局配置字典，包含 template_prompt 和 special_pages
    """
    data = json.loads(Path(template_json).read_text(encoding="utf-8"))

    # 读取全局配置
    global_config = {
        "template_prompt": data.get("template_prompt", {}),
        "special_pages": data.get("special_pages", {}),
    }

    allowed = None
    if template_list and Path(template_list).exists():
        content = Path(template_list).read_text(encoding="utf-8").strip()
        if content:
            allowed = {
                int(item.strip())
                for item in content.replace(",", " ").split()
                if item.strip().isdigit()
            }

    templates = {}
    manifest = {
        item["template_page_num"]: item
        for item in data.get("manifest", [])
        if "template_page_num" in item
    }

    for page in data.get("ppt_pages", []):
        num = page.get("template_page_num")
        if not isinstance(num, int):
            continue
        if allowed and num not in allowed:
            continue

        schema = page.get("content", {})
        fields = _collect_fields(schema)
        meta = page.get("meta", {}) or manifest.get(num, {})
        page_note = page.get("page_note") or meta.get("notes", "")

        templates[num] = {
            "page_type": page.get("page_type", f"模板第{num}页"),
            "schema": schema,
            "meta": meta,
            "page_note": page_note,
            "text_fields": [f for f in fields if not f["is_image"]],
            "image_fields": [f for f in fields if f["is_image"]],
        }

    if not templates:
        raise ValueError("模板列表为空，无法匹配。")

    return templates, global_config


def clone_schema(schema: Dict) -> Dict:
    """深拷贝 schema。"""
    return json.loads(json.dumps(schema, ensure_ascii=False))


def assign_in_schema(schema: Dict, path: List[str], value: str) -> None:
    """在 schema 中按路径赋值。"""
    node = schema
    for key in path[:-1]:
        node = node.setdefault(key, {})
    leaf = node.get(path[-1])
    if isinstance(leaf, dict) and "type" in leaf:
        leaf["value"] = value
    else:
        node[path[-1]] = value


def get_in_schema(content: Dict, path: List[str]) -> Optional[str]:
    """从嵌套字典中获取值。"""
    node = content
    for key in path:
        if isinstance(node, dict) and key in node:
            node = node[key]
        else:
            return None
    if isinstance(node, dict) and "value" in node:
        return node.get("value")
    return node if isinstance(node, str) else None


def empty_content(template_info: Dict) -> Dict:
    """创建空内容的 schema。"""
    result = clone_schema(template_info["schema"])
    for field in template_info["text_fields"]:
        assign_in_schema(result, list(field["path"]), "")
    for field in template_info["image_fields"]:
        assign_in_schema(result, list(field["path"]), "")
    return result


def strip_values(node: Any) -> Any:
    """递归移除 schema 中的 type 包装，只保留 value。"""
    if isinstance(node, dict):
        if "type" in node and "value" in node:
            return node.get("value", "")
        return {k: strip_values(v) for k, v in node.items()}
    if isinstance(node, list):
        return [strip_values(item) for item in node]
    return node
