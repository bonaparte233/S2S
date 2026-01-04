"""LLM Prompt 构建模块。"""

from __future__ import annotations

import base64
import mimetypes
import os
from typing import Any, Dict, List, Optional


def encode_image(image_path: str) -> Optional[str]:
    """读取图片文件并返回 base64 字符串。"""
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")
    except FileNotFoundError:
        print(f"⚠️ 警告：图片文件不存在：{image_path}")
        return None
    except Exception as e:
        print(f"⚠️ 警告：读取图片失败 {image_path}：{e}")
        return None


def _describe_fields(fields: List[Dict], is_image: bool = False) -> str:
    """详细描述每个字段。"""
    if not fields:
        return "（无）"
    lines = []
    for idx, field in enumerate(fields, 1):
        name = "/".join(field["path"])
        hint = field.get("hint") or ""
        max_chars = field.get("max_chars")
        required = field.get("required", False)

        parts = [f"  【字段{idx}】{name}"]
        if hint:
            parts.append(f"    → 填写要求：{hint}")
        if required:
            parts.append(f"    → ⚠️ 必填")
        else:
            parts.append(f"    → 可选（无匹配内容时留空）")
        if max_chars and not is_image:
            parts.append(f"    → 字数限制：≤{max_chars}字")

        lines.append("\n".join(parts))
    return "\n\n".join(lines)


def build_fill_prompt(
    template_info: Dict,
    raw_text: str,
    images: List[str],
    is_multimodal: bool = False,
    user_prompt: Optional[str] = None,
    metadata: Optional[Dict] = None,
    section_context: Optional[Dict[str, Optional[str]]] = None,
    global_config: Optional[Dict] = None,
) -> str:
    """构建 LLM 填充 prompt。

    Args:
        global_config: 全局配置，包含 template_prompt.fill_guide 等
    """
    text_desc = _describe_fields(template_info["text_fields"], is_image=False)
    image_desc = _describe_fields(template_info["image_fields"], is_image=True)

    if is_multimodal:
        image_section = f"已附加 {len(images)} 张图片"
    else:
        image_section = "无" if not images else "\n".join(images)

    page_note = template_info.get("page_note") or ""
    meta = template_info.get("meta") or {}
    if not page_note:
        page_note = meta.get("notes", "")

    page_note_section = ""
    if page_note:
        page_note_section = f"""
════════════════════════════════════════
📋 【本页特殊说明 - 必须优先遵守】
{page_note}
════════════════════════════════════════
"""

    multimodal_instruction = ""
    if is_multimodal and images:
        multimodal_instruction = f"""
🖼️ 图片处理说明：
已附加 {len(images)} 张图片。图片与文本的关系：
- 每张图片紧跟在相关文本段落的下方，是对上方文本的补充说明
- 讲稿中的 `[图片资源: ...]` 标记指示图片在原文中的位置
- 请根据图片字段的 hint 要求，将图片填入对应字段
- 如果 page_note 中有图片选择规则（如"选择1,3"），必须严格遵守
"""

    metadata_section = ""
    if metadata and any(metadata.values()):
        meta_items = []
        if metadata.get("course"):
            meta_items.append(
                f"课程名称：{metadata['course']}（填入 hint 中要求'课程'、'项目'的字段）"
            )
        if metadata.get("college"):
            meta_items.append(
                f"学院名称：{metadata['college']}（填入 hint 中要求'学院'、'单位'的字段）"
            )
        if metadata.get("lecturer"):
            meta_items.append(
                f"主讲人：{metadata['lecturer']}（填入 hint 中要求'主讲'、'讲师'、'姓名'的字段）"
            )
        if meta_items:
            # 从全局配置读取 fill_guide，否则使用默认
            template_prompt = (global_config or {}).get("template_prompt", {})
            fill_guide = template_prompt.get("fill_guide", "")
            if not fill_guide:
                fill_guide = "⚠️ 注意：主讲人姓名只能填入 hint 明确要求'主讲人'或'姓名'的字段，绝对不能填入'一级标题'、'章节标题'等标题类字段！"
            metadata_section = f"""
📌 预设信息（严格按照括号内的指示填入对应字段）：
{chr(10).join('- ' + item for item in meta_items)}
{fill_guide}
"""

    # 从全局配置读取 section_tracking 和 section_field_mappings
    template_prompt = (global_config or {}).get("template_prompt", {})
    section_tracking = template_prompt.get("section_tracking", True)
    section_field_mappings = template_prompt.get("section_field_mappings", {})

    context_section = ""
    if section_tracking and section_context and any(section_context.values()):
        ctx_items = []
        # 使用配置的映射或默认映射（只有两级）
        chapter_hint = section_field_mappings.get("chapter", "'章节'、'一级标题'")
        section_hint = section_field_mappings.get("section", "'知识点'、'二级标题'")

        if section_context.get("chapter"):
            ctx_items.append(f"章节：{section_context['chapter']}（填入 hint 中要求{chapter_hint}的字段）")
        if section_context.get("section"):
            ctx_items.append(f"知识点：{section_context['section']}（填入 hint 中要求{section_hint}的字段）")
        if ctx_items:
            context_section = f"""
📚 当前章节上下文（按照括号内的指示填入对应字段）：
{chr(10).join('- ' + item for item in ctx_items)}
"""

    prompt = f"""请根据讲稿内容填充模板《{template_info["page_type"]}》的各个字段。
{page_note_section}
═══════════════════════════════════════
📝 【字段定义 - 严格按照 hint 要求填写】

每个字段的 hint 是填写该字段的最重要依据，必须严格遵守。

【文本字段】
{text_desc}

【图片字段】
{image_desc}

可用图片：{image_section}
════════════════════════════════════════
{metadata_section}{context_section}{multimodal_instruction}
📌 填写原则：
1. 每个字段严格按照其 hint 要求填写，hint 是唯一的填写依据
2. 必填字段不得留空，可选字段无匹配内容时留空
3. 严格遵守字数限制，超出限制的内容需精简
4. 从讲稿中提炼要点，不要照搬原文
5. 如果 page_note 中有特殊规则（如图片选择规则），必须严格遵守

输出格式：
{{
  "texts": ["字段1内容", "字段2内容", ...],
  "images": ["图片路径1", "图片路径2", ...]
}}

数组顺序必须与上述字段定义顺序一致。

讲稿内容：
{raw_text}
"""
    if user_prompt:
        prompt += f"\n\n用户额外要求：\n{user_prompt}"

    return prompt


def build_multimodal_messages(
    template_info: Dict,
    raw_text: str,
    images: List[str],
    user_prompt: Optional[str] = None,
    metadata: Optional[Dict] = None,
    section_context: Optional[Dict[str, Optional[str]]] = None,
    global_config: Optional[Dict] = None,
) -> List[Dict]:
    """构建多模态消息。"""
    prompt_text = build_fill_prompt(
        template_info,
        raw_text,
        images,
        is_multimodal=True,
        user_prompt=user_prompt,
        metadata=metadata,
        section_context=section_context,
        global_config=global_config,
    )

    content: List[Dict[str, Any]] = [{"type": "text", "text": prompt_text}]

    for img_path in images:
        if not os.path.exists(img_path):
            print(f"⚠️ 警告：跳过不存在的图片：{img_path}")
            continue

        mime_type, _ = mimetypes.guess_type(img_path)
        if not mime_type:
            mime_type = "image/jpeg"

        base64_str = encode_image(img_path)
        if not base64_str:
            continue

        data_url = f"data:{mime_type};base64,{base64_str}"
        content.append({"type": "image_url", "image_url": {"url": data_url}})

    return [{"role": "user", "content": content}]
