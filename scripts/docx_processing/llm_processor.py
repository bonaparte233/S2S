"""LLM 调用逻辑模块。"""

from __future__ import annotations

import mimetypes
import os
from typing import Any, Dict, List, Optional

from scripts.llm_client import (
    BaseLLM,
    DeepSeekLLM,
    GLMLLM,
    LocalLLM,
    QwenVLLM,
    TaichuLLM,
)
from scripts.docx_processing.constants import DEBUG_LLM
from scripts.docx_processing.json_utils import ensure_json_array, ensure_json_object
from scripts.docx_processing.llm_prompts import (
    build_fill_prompt,
    build_multimodal_messages,
    encode_image,
)
from scripts.docx_processing.template_utils import (
    assign_in_schema,
    clone_schema,
)


def is_multimodal_llm(llm: Optional[BaseLLM]) -> bool:
    """检查是否为多模态模型。"""
    return isinstance(llm, (TaichuLLM, GLMLLM))


def choose_llm(
    enable: bool,
    provider: str,
    model: Optional[str],
    base_url: Optional[str] = None,
) -> Optional[BaseLLM]:
    """根据配置选择 LLM 实例。"""
    if not enable:
        return None
    provider = (provider or "").lower()

    if provider == "deepseek":
        return DeepSeekLLM(model=model or "deepseek-chat")
    if provider == "local":
        return LocalLLM(model=model)
    if provider == "qwen":
        endpoint = base_url or os.getenv("QWEN_VLLM_BASE_URL")
        if not endpoint:
            raise ValueError("Qwen provider 需要提供 base_url 或设置 QWEN_VLLM_BASE_URL。")
        return QwenVLLM(base_url=endpoint)
    if provider == "taichu":
        return TaichuLLM(model=model or "taichu4_vl_32b", base_url=base_url)
    if provider in ("glm", "zhipu"):
        return GLMLLM(model=model or "glm-4.5v", base_url=base_url)
    raise ValueError(f"暂不支持的大模型提供商：{provider}")


def _lookup_field_value(field: Dict, payload: Any, fallback_list: Any, idx: int) -> str:
    """从 LLM 响应中查找字段值。"""
    path = list(field["path"])
    key = "/".join(path)

    def normalize(text: str) -> str:
        return "".join(ch for ch in text.lower() if ch not in {" ", "_"})

    if isinstance(payload, dict):
        norm_targets = {normalize(key), normalize(path[-1])}
        for candidate_key, candidate_val in payload.items():
            norm = normalize(candidate_key)
            if norm in norm_targets or any(
                target and target in norm for target in norm_targets
            ):
                return candidate_val
        values = list(payload.values())
        if idx < len(values):
            return values[idx]
        return values[-1] if values else ""

    if isinstance(fallback_list, list):
        return fallback_list[idx] if idx < len(fallback_list) else ""
    return ""


def _simple_fill(template_info: Dict, raw_text: str, images: List[str]) -> Dict:
    """简单填充（无 LLM）。"""
    result = clone_schema(template_info["schema"])
    text_fields = template_info["text_fields"]
    image_fields = template_info["image_fields"]

    if text_fields:
        assign_in_schema(result, list(text_fields[0]["path"]), raw_text)
        for field in text_fields[1:]:
            assign_in_schema(result, list(field["path"]), "")

    for idx, field in enumerate(image_fields):
        value = images[idx] if idx < len(images) else ""
        assign_in_schema(result, list(field["path"]), value)

    return result


def llm_fill_slide(
    llm: BaseLLM,
    template_info: Dict,
    raw_text: str,
    images: List[str],
    user_prompt: Optional[str] = None,
    use_multimodal: bool = True,
    metadata: Optional[Dict] = None,
    section_context: Optional[Dict[str, Optional[str]]] = None,
    global_config: Optional[Dict] = None,
) -> Dict:
    """使用 LLM 填充单页幻灯片内容。"""
    if not llm:
        return _simple_fill(template_info, raw_text, images)

    if use_multimodal and is_multimodal_llm(llm) and images:
        messages = build_multimodal_messages(
            template_info, raw_text, images, user_prompt, metadata, section_context,
            global_config,
        )
    else:
        prompt = build_fill_prompt(
            template_info, raw_text, images,
            user_prompt=user_prompt, metadata=metadata, section_context=section_context,
            global_config=global_config,
        )
        messages = [{"role": "user", "content": prompt}]

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("🔍 [DEBUG] LLM 请求 (llm_fill_slide)")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.2)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("📥 [DEBUG] LLM 响应 (llm_fill_slide)")
        print(f"{'=' * 60}")
        print(f"{response[:500]}...")
        print(f"{'=' * 60}\n")

    try:
        data = ensure_json_object(response)
        texts = data.get("texts", [])
        imgs = data.get("images", [])
    except Exception:
        texts, imgs = [raw_text], images

    result = clone_schema(template_info["schema"])
    text_fields = template_info["text_fields"]
    image_fields = template_info["image_fields"]

    for idx, field in enumerate(text_fields):
        value = _lookup_field_value(
            field, texts, texts if isinstance(texts, list) else None, idx
        )
        assign_in_schema(result, list(field["path"]), value)

    for idx, field in enumerate(image_fields):
        value = _lookup_field_value(
            field, imgs, imgs if isinstance(imgs, list) else None, idx
        )
        if not value and isinstance(images, list) and idx < len(images):
            value = images[idx]
        assign_in_schema(result, list(field["path"]), value)

    return result


def llm_plan_slides(
    llm: BaseLLM,
    doc_text: str,
    templates: Dict[int, Dict],
    images: List[str],
    user_prompt: Optional[str] = None,
) -> List[Dict]:
    """使用 LLM 规划幻灯片。"""
    template_desc = "\n".join(
        f"- 模板 {info['page_type']} (编号 {num}): "
        f"文本{len(info['text_fields'])}项, 图片{len(info['image_fields'])}项"
        for num, info in templates.items()
    )

    if is_multimodal_llm(llm) and images:
        image_section = f"已附加 {len(images)} 张图片供你参考"
    else:
        image_section = "无" if not images else "\n".join(images)

    multimodal_instruction = ""
    if is_multimodal_llm(llm) and images:
        multimodal_instruction = f"""
🖼️ 多模态图片理解（重要）：
我已附带了 {len(images)} 张图片，这些图片是讲稿的重要组成部分，请务必认真处理。
"""

    prompt = f"""
请将以下讲稿拆分成若干张 PPT，每张幻灯片选择一个模板，并输出 JSON 数组，每个元素包含：
- template_page_num: 模板编号
- page_type: 模板名称
- texts: 按模板文本字段顺序给出的内容数组
- images: 按模板图片字段顺序给出的内容数组

模板信息：
{template_desc}

可用图片路径：
{image_section}

{multimodal_instruction}

📌 核心原则：幻灯片是讲稿的精炼要点，而非照搬全文。

讲稿全文：
{doc_text}
"""
    if user_prompt:
        prompt += f"\n\n用户额外要求：\n{user_prompt}"

    if is_multimodal_llm(llm) and images:
        content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
        for img_path in images:
            if not os.path.exists(img_path):
                continue
            mime_type, _ = mimetypes.guess_type(img_path)
            if not mime_type:
                mime_type = "image/jpeg"
            base64_str = encode_image(img_path)
            if not base64_str:
                continue
            data_url = f"data:{mime_type};base64,{base64_str}"
            content.append({"type": "image_url", "image_url": {"url": data_url}})
        messages = [{"role": "user", "content": content}]
    else:
        messages = [{"role": "user", "content": prompt}]

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("🔍 [DEBUG] LLM 请求 (llm_plan_slides)")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.3)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("📥 [DEBUG] LLM 响应 (llm_plan_slides)")
        print(f"{'=' * 60}")
        print(f"{response[:500]}...")
        print(f"{'=' * 60}\n")

    try:
        return ensure_json_array(response)
    except Exception:
        raise ValueError("模型输出无法解析为 JSON 数组，请检查提示或重试。")


def llm_preprocess_script(
    llm: BaseLLM,
    doc_text: str,
    templates: Dict[int, Dict],
    images: List[str],
    user_prompt: Optional[str] = None,
    has_marker: bool = False,
    global_config: Optional[Dict] = None,
) -> str:
    """使用 LLM 预处理讲稿。

    Args:
        llm: LLM 实例
        doc_text: 原始讲稿文本
        templates: 模板定义字典
        images: 图片路径列表
        user_prompt: 用户自定义提示
        has_marker: 讲稿是否已有【PPT】标记
        global_config: 全局配置，包含 template_prompt 等
    """
    if not llm:
        raise ValueError("预处理讲稿需要启用 LLM。")

    # 从全局配置读取 preprocess_guide
    template_prompt = (global_config or {}).get("template_prompt", {})
    preprocess_guide = template_prompt.get("preprocess_guide", "")

    # 构建模板描述
    template_desc_lines = []
    for num, info in templates.items():
        text_count = len(info["text_fields"])
        image_count = len(info["image_fields"])
        page_type = info["page_type"]
        has_required_image = any(f.get("required") for f in info["image_fields"])

        page_note = info.get("page_note") or ""
        meta = info.get("meta") or {}
        if not page_note:
            page_note = meta.get("notes", "")

        text_fields_desc = []
        for field in info["text_fields"]:
            path = field.get("path", ())
            name = "/".join(path) if path else "未命名"
            hint = field.get("hint", "")
            max_chars = field.get("max_chars")
            required = field.get("required", False)
            req_mark = "必填" if required else "可选"
            char_limit = f"≤{max_chars}字" if max_chars else ""
            desc_parts = [name]
            if hint:
                desc_parts.append(f"({hint})")
            if char_limit:
                desc_parts.append(f"[{char_limit}]")
            desc_parts.append(f"[{req_mark}]")
            text_fields_desc.append(f"    - {' '.join(desc_parts)}")

        image_fields_desc = []
        for field in info["image_fields"]:
            path = field.get("path", ())
            name = "/".join(path) if path else "图片"
            hint = field.get("hint", "")
            required = "必填" if field.get("required") else "可选"
            desc_parts = [name]
            if hint:
                desc_parts.append(f"({hint})")
            desc_parts.append(f"[{required}]")
            image_fields_desc.append(f"    - {' '.join(desc_parts)}")

        desc = f"【PPT{num}】{page_type}"
        if has_required_image:
            desc += " ⚠️需要图片"
        if page_note:
            desc += f"\n  📋 特殊说明：{page_note}"
        desc += f"\n  文本字段({text_count}个):"
        if text_fields_desc:
            desc += "\n" + "\n".join(text_fields_desc)
        else:
            desc += "\n    （无）"
        if image_count > 0:
            desc += f"\n  图片字段({image_count}个):"
            desc += "\n" + "\n".join(image_fields_desc)
        template_desc_lines.append(desc)

    template_desc = "\n\n".join(template_desc_lines)

    # 根据是否已有标记，构建不同的 prompt
    if has_marker:
        # 已有标记：保持分页，只优化文本
        prompt = f"""你是一位专业的演讲稿编辑。讲稿已经分好页了，请**保持原有分页结构**，只优化文本表达。

## 任务说明

1. **保持分页不变**：讲稿中的【PPT编号】标记表示分页，必须原样保留，不能修改、删除或重新分配
2. **优化文本表达**：将口语化表达改为正式书面语，但不改变语义
3. **保持内容完整**：不要精简或删减内容，保持原文语义完整

⚠️ **重要**：分页已由用户指定，你只能优化文本，不能改变分页！

## 可用模板（仅供参考）

{template_desc}

## 注意事项

1. 【PPT编号】标记必须原样保留，不能修改
2. 不要精简内容，保持讲稿原文的完整性
3. 保留所有图片引用 `[图片资源: ...]`，位置不变
4. 不要输出任何解释说明，只输出优化后的讲稿

## 原始讲稿

{doc_text}
"""
    else:
        # 无标记：自动分页
        if images:
            image_info = f"讲稿中包含 {len(images)} 张图片，请在适当位置保留图片引用。"
        else:
            image_info = "⚠️ 讲稿中【没有图片】！"
            templates_need_image = [
                num for num, info in templates.items()
                if any(f.get("required") for f in info["image_fields"])
            ]
            if templates_need_image:
                image_info += f"\n【禁止】使用以下需要图片的模板：{templates_need_image}"

        # 构建可选的 preprocess_guide 部分
        preprocess_guide_section = ""
        if preprocess_guide:
            preprocess_guide_section = f"\n{preprocess_guide}\n"

        prompt = f"""你是一位专业的演讲稿编辑。请将以下原始讲稿**分页**，为每个部分选择合适的 PPT 模板。

## 任务说明

1. **分析讲稿结构**：理解讲稿的主题、逻辑和内容层次
2. **选择合适模板**：根据内容为每个部分选择最合适的 PPT 模板
3. **添加页码标记**：在每个部分开头用【PPT编号】标记该部分使用的模板
4. **保持内容完整**：不要精简或删减内容，保持原文语义完整，只需将口语化表达改为正式书面语

⚠️ **重要**：这一步只负责**分页和选择模板**，不要精简内容！

## 可用模板

{template_desc}
{preprocess_guide_section}
## 图片处理规则

{image_info}

**图片与模板匹配规则**：
- 图片标记 `[图片资源: xxx]` 必须与其**紧邻的上方文本**分到同一页
- 如果某页内容包含图片，应选择带图片字段的模板
- 如果某页内容**没有图片**，**禁止**选择标有"⚠️需要图片"的模板
- 图片是对上方文本的说明或示例，不能与文本分离

## 输出格式要求

输出为 Markdown 格式，每个 PPT 页面以【PPT编号】开头，例如：

```
【PPT2】
# 课程介绍

本课程将带您了解人工智能的基础知识，涵盖机器学习、深度学习等核心领域。

【PPT5】
# 机器学习基础

机器学习是人工智能的核心技术，它让计算机能够从数据中学习规律。

[图片资源: doc_image_1.png]
```

## 注意事项

1. 每个【PPT编号】标记必须独占一行
2. 编号必须是上面模板列表中存在的编号
3. **没有图片的页面禁止使用需要图片的模板**
4. 不要精简内容，保持讲稿原文的完整性
5. 图片引用 `[图片资源: ...]` 必须与上方文本保持在同一页
6. 不要输出任何解释说明，只输出分页后的讲稿

## 原始讲稿

{doc_text}
"""

    if user_prompt:
        prompt += f"\n\n## 用户额外要求\n\n{user_prompt}"

    if is_multimodal_llm(llm) and images:
        content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
        for img_path in images:
            if not os.path.exists(img_path):
                continue
            mime_type, _ = mimetypes.guess_type(img_path)
            if not mime_type:
                mime_type = "image/jpeg"
            base64_str = encode_image(img_path)
            if not base64_str:
                continue
            data_url = f"data:{mime_type};base64,{base64_str}"
            content.append({"type": "image_url", "image_url": {"url": data_url}})
        messages = [{"role": "user", "content": content}]
    else:
        messages = [{"role": "user", "content": prompt}]

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("🔍 [DEBUG] LLM 请求 (llm_preprocess_script)")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.3)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("📥 [DEBUG] LLM 响应 (llm_preprocess_script)")
        print(f"{'=' * 60}")
        print(f"{response[:1000]}...")
        print(f"{'=' * 60}\n")

    result = response.strip()
    if result.startswith("```"):
        lines = result.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        result = "\n".join(lines)

    return result
