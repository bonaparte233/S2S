"""根据 DOCX 讲稿和模板定义生成 JSON，可选调用 DeepSeek LLM。"""

from __future__ import annotations

import argparse
import json
import re
import secrets
import shutil
from datetime import datetime
from pathlib import Path
import os
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn

from scripts.llm_client import (
    BaseLLM,
    DeepSeekLLM,
    GLMLLM,
    LocalLLM,
    QwenVLLM,
    TaichuLLM,
)
import base64
import mimetypes

MARKER_RE = re.compile(r"【PPT(\d+)】")
IMAGE_NAME_TEMPLATE = "doc_image_{idx}.{ext}"

# 调试标志：设置为 True 时打印 LLM 请求和响应
DEBUG_LLM = os.getenv("DEBUG_LLM", "false").lower() in ("true", "1", "yes")


def _create_run_dir(base_dir: Path = Path("temp")) -> Path:
    """创建带时间戳前缀的运行目录，方便前端一次处理对应到单独目录。"""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = secrets.token_hex(2)
    run_dir = base_dir / f"script-{timestamp}-{suffix}"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def parse_docx_blocks(doc_path: str, image_dir: Path) -> Tuple[List[Dict], bool, Dict]:
    """读取 DOCX 并按照 PPT 标记拆分内容，同时保存提取的图片与课程元信息。"""
    doc = Document(doc_path)
    image_dir.mkdir(parents=True, exist_ok=True)
    slides: List[Dict] = []
    current: Optional[Dict] = None
    buffer: List[str] = []
    has_marker = False
    image_counter = 1
    metadata: Dict[str, str] = {}

    def flush():
        nonlocal current, buffer
        if current is None and buffer:
            slides.append(
                {"template_hint": None, "text": "\n".join(buffer).strip(), "images": []}
            )
        elif current:
            current["text"] = "\n".join(buffer).strip()
            slides.append(current)
        buffer = []
        current = None

    def ensure_block():
        nonlocal current
        if current is None:
            current = {"template_hint": None, "text": "", "images": []}

    def attach_images(paths: List[str]):
        if not paths:
            return
        ensure_block()
        current.setdefault("images", []).extend(paths)

    def extract_images(paragraph) -> List[str]:
        nonlocal image_counter
        images = []
        for element in paragraph._p.iter():
            if element.tag not in {qn("a:blip"), qn("pic:blip")}:
                continue
            r_id = element.get(qn("r:embed"))
            if not r_id:
                continue
            part = paragraph.part.related_parts.get(r_id)
            if not part:
                continue
            ext = part.filename.split(".")[-1].lower() or "png"
            filename = image_dir / IMAGE_NAME_TEMPLATE.format(
                idx=image_counter, ext=ext
            )
            with open(filename, "wb") as f:
                f.write(part.blob)
            images.append(str(filename))
            image_counter += 1
        return images

    for para in doc.paragraphs:
        text = para.text
        stripped = text.strip()
        meta_match = re.match(r"^(课程名称|学院名称|主讲教师)[：:]\s*(.+)$", stripped)
        if meta_match:
            key = meta_match.group(1)
            value = meta_match.group(2).strip()
            if key == "课程名称":
                metadata["course"] = value
            elif key == "学院名称":
                metadata["college"] = value
            elif key == "主讲教师":
                metadata["lecturer"] = value
            continue
        images = extract_images(para)
        if images:
            attach_images(images)
            for img_path in images:
                buffer.append(f"[图片资源: {img_path}]")

        matches = list(MARKER_RE.finditer(text))
        if matches:
            idx = 0
            for match in matches:
                prefix = text[idx : match.start()].strip()
                if prefix:
                    buffer.append(prefix)
                flush()
                has_marker = True
                current = {
                    "template_hint": int(match.group(1)),
                    "text": "",
                    "images": [],
                }
                idx = match.end()
            remainder = text[idx:].strip()
            if remainder:
                buffer.append(remainder)
            continue

        if stripped:
            buffer.append(stripped)

    flush()
    if not slides:
        full_text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
        slides.append({"template_hint": None, "text": full_text, "images": []})
    return slides, has_marker, metadata


def load_template_defs(template_json: str, template_list: str) -> Dict[int, Dict]:
    data = json.loads(Path(template_json).read_text(encoding="utf-8"))
    allowed = None
    if template_list and Path(template_list).exists():
        allowed = {
            int(item.strip())
            for item in Path(template_list)
            .read_text(encoding="utf-8")
            .replace(",", " ")
            .split()
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
        templates[num] = {
            "page_type": page.get("page_type", f"模板第{num}页"),
            "schema": schema,
            "meta": page.get("meta", {}) or manifest.get(num, {}),
            "text_fields": [field for field in fields if not field["is_image"]],
            "image_fields": [field for field in fields if field["is_image"]],
        }
    if not templates:
        raise ValueError("模板列表为空，无法匹配。")
    return templates


def _collect_fields(schema, prefix=None):
    results = []
    prefix = prefix or ()
    if isinstance(schema, dict):
        if "type" in schema and "value" in schema:
            field_type = schema.get("type") or "text"
            results.append(
                {
                    "path": prefix,
                    "is_image": field_type.lower() == "image",
                    "hint": schema.get("hint") or "",
                    "max_chars": schema.get("max_chars"),
                    "required": bool(schema.get("required", False)),
                }
            )
        else:
            for key, value in schema.items():
                results.extend(_collect_fields(value, prefix + (key,)))
    else:
        # 兼容旧格式：字符串或其他类型视为文本叶子
        results.append(
            {
                "path": prefix,
                "is_image": any(
                    "图片" in seg or "image" in seg.lower() for seg in prefix
                ),
                "hint": "",
                "max_chars": None,
                "required": False,
            }
        )
    return results


def _clone_schema(schema: Dict) -> Dict:
    return json.loads(json.dumps(schema, ensure_ascii=False))


def _assign_in_schema(schema: Dict, path: List[str], value: str):
    node = schema
    for key in path[:-1]:
        node = node.setdefault(key, {})
    leaf = node.get(path[-1])
    if isinstance(leaf, dict) and "type" in leaf:
        leaf["value"] = value
    else:
        node[path[-1]] = value


def _is_multimodal_llm(llm: Optional[BaseLLM]) -> bool:
    """检查是否为多模态模型"""
    return isinstance(llm, (TaichuLLM, GLMLLM))


def _simple_fill(template_info: Dict, raw_text: str, images: List[str]) -> Dict:
    result = _clone_schema(template_info["schema"])
    text_fields = template_info["text_fields"]
    image_fields = template_info["image_fields"]
    if text_fields:
        _assign_in_schema(result, list(text_fields[0]["path"]), raw_text)
        for field in text_fields[1:]:
            _assign_in_schema(result, list(field["path"]), "")
    for idx, field in enumerate(image_fields):
        value = images[idx] if idx < len(images) else ""
        _assign_in_schema(result, list(field["path"]), value)
    return result


def _build_prompt(
    template_info: Dict,
    raw_text: str,
    images: List[str],
    is_multimodal: bool = False,
    user_prompt: Optional[str] = None,
) -> str:
    def describe_fields(fields):
        if not fields:
            return "无"
        lines = []
        for idx, field in enumerate(fields, 1):
            name = "/".join(field["path"])
            hint = field.get("hint") or "填写内容"
            extra = []
            if field.get("max_chars"):
                extra.append(f"≤{field['max_chars']}字")
            if field.get("required"):
                extra.append("必填")
            extra_note = f"（{'，'.join(extra)}）" if extra else ""
            lines.append(f"{idx}. {name}：{hint}{extra_note}")
        return "\n".join(lines)

    text_desc = describe_fields(template_info["text_fields"])
    image_desc = describe_fields(template_info["image_fields"])

    # 对于多模态模型，不需要列出图片路径（图片已通过 base64 附加）
    if is_multimodal:
        image_section = f"已附加 {len(images)} 张图片供你参考"
    else:
        image_section = "无" if not images else "\n".join(images)

    meta = template_info.get("meta") or {}
    scene = "、".join(meta.get("scene", [])) or "通用"
    layout = meta.get("layout", template_info["page_type"])
    style = meta.get("style", "")
    note = meta.get("notes", "")

    multimodal_instruction = ""
    if is_multimodal and images:
        multimodal_instruction = f"""
🖼️ 多模态图片理解（重要）：
我已附带了 {len(images)} 张图片，这些图片是讲稿的重要组成部分，请务必认真处理。

图片与文本的关系：
- 每张图片都紧跟在相关文本段落的下方（图片在文本下方）
- 图片是对上方文本的补充说明、示例或可视化
- 讲稿文本中的 `[图片资源: ...]` 标记仅用于指示图片在原文中的位置

你的任务：
1. 仔细查看每张图片的内容，理解图片传达的信息
2. 分析图片与上下文文本的关系，确定图片所属的主题
3. 将图片放入合适的图片字段（通常与相关文本在同一页 PPT）
4. 在 JSON 的 "images" 数组对应位置填入该图片的完整路径
5. 根据图片内容优化文本描述，使其更准确、更生动
6. 如果图片字段不需要图片，请留空字符串

⚠️ 重要：不要忽略图片！图片是讲稿的核心内容之一，必须合理使用。
"""

    prompt = f"""
请阅读以下讲稿并生成一个 JSON，对模板《{template_info["page_type"]}》的文本/图片字段进行填充。
模板布局：{layout}；使用场景：{scene}；风格提示：{style}
注意事项：{note}
{multimodal_instruction}

📌 核心原则（幻灯片 vs 讲稿）：
讲稿是演讲者手里的稿子，是他演讲时要说的完整内容。
幻灯片是投影给观众看的，应该是讲稿的**精炼要点**，而非照搬全文。
你需要把讲稿内容**提炼、概括、分点**后放到幻灯片上。

✅ 正确做法：
- 提取核心观点，用简洁的短语或短句表达
- 使用要点列表（如"1. xxx  2. xxx"或"• xxx"）
- 删除口语化表达、过渡语、详细解释
- 保留关键数据、专有名词、核心结论

❌ 错误做法：
- 把讲稿的长段落直接复制到幻灯片
- 保留"接下来我们来看""正如前面所说"等口语
- 内容过于详细，像在读文章

⚠️ 严格要求（必须遵守）：
1. 所有标记为"required"的字段必须填写，绝对不得留空。
2. 每个文本字段都有字数上限（max_chars），你生成的内容绝对不能超过这个限制。
3. 内容必须精炼！把讲稿的完整表述压缩为幻灯片要点，只保留核心信息。
4. 违反字数限制的输出将被视为无效，必须重新生成。
5. 务必记住讲稿中提到的主讲人姓名、课程/讲座/项目名称等关键专有名词，并在所有需要这些信息的字段保持完全一致，不要改写。

该模板包含如下文本字段（按照顺序对应）：
{text_desc}

图片字段（若无可留空）：
{image_desc}

可用图片路径：
{image_section}

输出格式示例：
{{
  "texts": ["文本1", "文本2", "..."],
  "images": ["图片路径1", "图片路径2", "..."]
}}

请严格保持数组长度与字段数量一致，texts[1] 必须对应上述列表中的第 1 个文本字段，依此类推。
讲稿内容：
{raw_text}
"""
    # Append user prompt if provided
    if user_prompt:
        prompt += f"\n\n用户额外要求：\n{user_prompt}"

    return prompt


def _encode_image(image_path: str) -> Optional[str]:
    """Read image file and return base64 string.

    Returns:
        Base64 encoded string, or None if encoding fails.
    """
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")
    except FileNotFoundError:
        print(f"⚠️ 警告：图片文件不存在：{image_path}")
        return None
    except Exception as e:
        print(f"⚠️ 警告：读取图片失败 {image_path}：{e}")
        return None


def _build_multimodal_messages(
    template_info: Dict,
    raw_text: str,
    images: List[str],
    user_prompt: Optional[str] = None,
) -> List[Dict]:
    """构建多模态消息，用于 Taichu-VL 或 GLMV。"""
    prompt_text = _build_prompt(
        template_info, raw_text, images, is_multimodal=True, user_prompt=user_prompt
    )

    content: List[Dict[str, Any]] = [{"type": "text", "text": prompt_text}]

    # 添加图片
    for img_path in images:
        if not os.path.exists(img_path):
            print(f"⚠️ 警告：跳过不存在的图片：{img_path}")
            continue

        # Taichu-VL 使用 OpenAI 兼容格式，支持 data URL (base64)
        # 参考：https://docs.wair.ac.cn/intelligent/maas/visioIntro.html
        mime_type, _ = mimetypes.guess_type(img_path)
        if not mime_type:
            mime_type = "image/jpeg"

        base64_str = _encode_image(img_path)
        if not base64_str:  # 编码失败，跳过此图片
            continue

        data_url = f"data:{mime_type};base64,{base64_str}"

        content.append({"type": "image_url", "image_url": {"url": data_url}})

    return [{"role": "user", "content": content}]


def _lookup_field_value(field, payload, fallback_list, idx):
    path = list(field["path"])
    key = "/".join(path)

    def normalize(text):
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


def llm_fill_slide(
    llm: BaseLLM,
    template_info: Dict,
    raw_text: str,
    images: List[str],
    user_prompt: Optional[str] = None,
    use_multimodal: bool = True,
) -> Dict:
    """
    使用 LLM 填充单页幻灯片内容。

    Args:
        llm: LLM 实例
        template_info: 模板信息
        raw_text: 原始文本
        images: 图片路径列表
        user_prompt: 用户自定义 prompt
        use_multimodal: 是否使用多模态消息（默认 True）
                       当讲稿有 PPT 标记时，图片位置已确定，可设为 False
    """
    if not llm:
        return _simple_fill(template_info, raw_text, images)

    # 只有在允许使用多模态且模型支持多模态且有图片时，才使用多模态消息
    if use_multimodal and _is_multimodal_llm(llm) and images:
        messages = _build_multimodal_messages(
            template_info, raw_text, images, user_prompt
        )
    else:
        prompt = _build_prompt(template_info, raw_text, images, user_prompt=user_prompt)
        messages = [{"role": "user", "content": prompt}]

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print(f"🔍 [DEBUG] LLM 请求 (llm_fill_slide)")
        print(f"{'=' * 60}")
        # 检查实际发送的消息类型
        is_multimodal_message = messages and isinstance(
            messages[0].get("content"), list
        )
        if is_multimodal_message:
            print(f"📝 多模态消息 (文本 + {len(images)} 张图片)")
            # 只打印文本部分，图片太长不打印
            for msg in messages:
                if isinstance(msg.get("content"), list):
                    for item in msg["content"]:
                        if item.get("type") == "text":
                            print(f"文本内容:\n{item['text'][:500]}...")
        else:
            print(f"📝 文本消息:\n{messages[0]['content'][:500]}...")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.2)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print(f"📥 [DEBUG] LLM 响应 (llm_fill_slide)")
        print(f"{'=' * 60}")
        print(f"{response[:500]}...")
        print(f"{'=' * 60}\n")
    try:
        data = _ensure_json_object(response)
        texts = data.get("texts", [])
        imgs = data.get("images", [])
    except Exception:
        texts, imgs = [raw_text], images

    result = _clone_schema(template_info["schema"])
    text_fields = template_info["text_fields"]
    image_fields = template_info["image_fields"]

    for idx, field in enumerate(text_fields):
        value = _lookup_field_value(
            field, texts, texts if isinstance(texts, list) else None, idx
        )
        _assign_in_schema(result, list(field["path"]), value)

    for idx, field in enumerate(image_fields):
        value = _lookup_field_value(
            field, imgs, imgs if isinstance(imgs, list) else None, idx
        )
        if not value and isinstance(images, list) and idx < len(images):
            value = images[idx]
        _assign_in_schema(result, list(field["path"]), value)

    return result


def llm_plan_slides(
    llm: BaseLLM,
    doc_text: str,
    templates: Dict[int, Dict],
    images: List[str],
    user_prompt: Optional[str] = None,
) -> List[Dict]:
    template_desc = "\n".join(
        f"- 模板 {info['page_type']} (编号 {num}): 文本{len(info['text_fields'])}项, 图片{len(info['image_fields'])}项"
        for num, info in templates.items()
    )

    # 对于多模态模型，不需要列出图片路径（图片已通过 base64 附加）
    if _is_multimodal_llm(llm) and images:
        image_section = f"已附加 {len(images)} 张图片供你参考"
    else:
        image_section = "无" if not images else "\n".join(images)

    multimodal_instruction = ""
    if _is_multimodal_llm(llm) and images:
        multimodal_instruction = f"""
🖼️ 多模态图片理解（重要）：
我已附带了 {len(images)} 张图片，这些图片是讲稿的重要组成部分，请务必认真处理。

图片与文本的关系：
- 每张图片都紧跟在相关文本段落的下方（图片在文本下方）
- 图片是对上方文本的补充说明、示例或可视化
- 讲稿文本中的 `[图片资源: ...]` 标记仅用于指示图片在原文中的位置

你的任务：
1. 仔细查看每张图片的内容，理解图片传达的信息
2. 分析图片与上下文文本的关系，确定图片所属的主题
3. 将图片放入合适的模板的图片字段（通常与相关文本在同一页 PPT）
4. 在输出的 JSON 对象中，"images" 数组应包含你选择使用的图片完整路径
5. 根据图片内容优化文本描述，使其更准确、更生动

⚠️ 重要：不要忽略图片！图片是讲稿的核心内容之一，必须合理使用。
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

输出格式示例：
[
  {{
    "template_page_num": 4,
    "page_type": "目录页",
    "texts": ["目录标题", "条目1", "条目1说明", "..."],
    "images": [""]
  }},
  ...
]


📌 核心原则（幻灯片 vs 讲稿）：
讲稿是演讲者手里的稿子，是他演讲时要说的完整内容。
幻灯片是投影给观众看的，应该是讲稿的**精炼要点**，而非照搬全文。
你需要把讲稿内容**提炼、概括、分点**后放到幻灯片上。

✅ 正确做法：
- 提取核心观点，用简洁的短语或短句表达
- 使用要点列表（如"1. xxx  2. xxx"或"• xxx"）
- 删除口语化表达、过渡语、详细解释
- 保留关键数据、专有名词、核心结论

❌ 错误做法：
- 把讲稿的长段落直接复制到幻灯片
- 保留"接下来我们来看""正如前面所说"等口语
- 内容过于详细，像在读文章

⚠️ 严格要求（必须遵守）：
1. 所有标记为"required"的字段必须填写，绝对不得留空。
2. 每个文本字段都有字数上限（max_chars），你生成的内容绝对不能超过这个限制。
3. 内容必须精炼！把讲稿的完整表述压缩为幻灯片要点，只保留核心信息。
4. 违反字数限制的输出将被视为无效，必须重新生成。
5. 务必记住并重复使用讲稿中的主讲人姓名、课程/讲座/项目名称等关键专有名词，确保在所有幻灯片中需要填写专有名词的位置保持一致，不要随意改写或另造新名称。

讲稿全文：
{doc_text}
"""
    # Append user prompt if provided
    if user_prompt:
        prompt += f"\n\n用户额外要求：\n{user_prompt}"

    if _is_multimodal_llm(llm) and images:
        content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
        for img_path in images:
            if not os.path.exists(img_path):
                print(f"⚠️ 警告：跳过不存在的图片：{img_path}")
                continue
            mime_type, _ = mimetypes.guess_type(img_path)
            if not mime_type:
                mime_type = "image/jpeg"
            base64_str = _encode_image(img_path)
            if not base64_str:  # 编码失败，跳过此图片
                continue
            data_url = f"data:{mime_type};base64,{base64_str}"
            content.append({"type": "image_url", "image_url": {"url": data_url}})
        messages = [{"role": "user", "content": content}]
    else:
        messages = [{"role": "user", "content": prompt}]

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("🔍 [DEBUG] LLM 请求 (llm_plan_slides)")
        print(f"{'=' * 60}")
        if _is_multimodal_llm(llm) and images:
            print(f"📝 多模态消息 (文本 + {len(images)} 张图片)")
            # 只打印文本部分
            for msg in messages:
                if isinstance(msg.get("content"), list):
                    for item in msg["content"]:
                        if isinstance(item, dict) and item.get("type") == "text":
                            print(f"文本内容:\n{item['text'][:500]}...")
        else:
            print(f"📝 文本消息:\n{messages[0]['content'][:500]}...")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.3)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("📥 [DEBUG] LLM 响应 (llm_plan_slides)")
        print(f"{'=' * 60}")
        print(f"{response[:500]}...")
        print(f"{'=' * 60}\n")

    try:
        plan = _ensure_json_array(response)
        return plan
    except Exception:
        raise ValueError("模型输出无法解析为 JSON 数组，请检查提示或重试。")


def llm_preprocess_script(
    llm: BaseLLM,
    doc_text: str,
    templates: Dict[int, Dict],
    images: List[str],
    user_prompt: Optional[str] = None,
) -> str:
    """
    使用 LLM 将原始讲稿预处理为带【PPT】标记的中间讲稿。

    Args:
        llm: LLM 实例
        doc_text: 原始讲稿文本
        templates: 模板定义字典
        images: 图片路径列表
        user_prompt: 用户自定义提示

    Returns:
        带有【PPT1】【PPT2】等标记的 Markdown 格式讲稿
    """
    if not llm:
        raise ValueError("预处理讲稿需要启用 LLM。")

    # 构建模板描述
    template_desc_lines = []
    for num, info in templates.items():
        text_count = len(info["text_fields"])
        image_count = len(info["image_fields"])
        page_type = info["page_type"]

        # 获取文本字段的详细信息
        text_fields_desc = []
        for field in info["text_fields"]:
            name = field.get("name", "未命名")
            max_chars = field.get("max_chars", "无限制")
            text_fields_desc.append(f"    - {name}（最多{max_chars}字）")

        image_fields_desc = []
        for field in info["image_fields"]:
            name = field.get("name", "图片")
            image_fields_desc.append(f"    - {name}")

        desc = f"【PPT{num}】{page_type}\n  文本字段({text_count}个):\n"
        desc += "\n".join(text_fields_desc) if text_fields_desc else "    （无）"
        if image_count > 0:
            desc += f"\n  图片字段({image_count}个):\n"
            desc += "\n".join(image_fields_desc)
        template_desc_lines.append(desc)

    template_desc = "\n\n".join(template_desc_lines)

    # 图片信息和模板限制
    if images:
        image_info = f"讲稿中包含 {len(images)} 张图片，请在适当位置保留图片引用。"
        # 所有模板都可用
        available_templates = list(templates.keys())
    else:
        image_info = "⚠️ 讲稿中【没有图片】，请【只选择不包含图片字段的模板】！"
        # 只保留没有图片字段的模板
        available_templates = [
            num for num, info in templates.items() if len(info["image_fields"]) == 0
        ]
        image_info += f"\n可用模板编号：{available_templates}"

    prompt = f"""你是一位专业的演讲稿编辑。请将以下原始讲稿改写为适合 PPT 演示的正式演讲稿。

## 任务说明

1. **分析讲稿结构**：理解讲稿的主题、逻辑和内容层次
2. **选择合适模板**：根据内容为每个部分选择最合适的 PPT 模板
3. **添加页码标记**：在每个部分开头用【PPT编号】标记该部分使用的模板
4. **优化表达**：将内容改写为正式、简洁的演讲风格，但不改变原意
5. **控制篇幅**：根据每个模板的字数限制，精简内容使其适合 PPT 展示

## 可用模板

{template_desc}

## 图片信息

{image_info}

## 输出格式要求

输出为 Markdown 格式，每个 PPT 页面以【PPT编号】开头，例如：

```
【PPT2】
# 课程介绍

本课程将带您了解人工智能的基础知识...

【PPT4】
# 课程目录

1. 机器学习基础
2. 深度学习入门
3. 实践案例分析

【PPT5】
# 机器学习基础

机器学习是人工智能的核心技术...

[图片资源: doc_image_1.png]
```

## 注意事项

1. 每个【PPT编号】标记必须独占一行
2. 编号必须是上面模板列表中存在的编号
3. **重要**：如果讲稿没有图片，则【禁止】使用带图片字段的模板！
4. 内容要精炼，适合 PPT 展示，避免大段文字
5. 保留讲稿中的关键信息、专有名词和数据
6. 如有图片引用（[图片资源: ...]），请保留在合适的位置
7. 不要输出任何解释说明，只输出改写后的讲稿

## 原始讲稿

{doc_text}
"""

    if user_prompt:
        prompt += f"\n\n## 用户额外要求\n\n{user_prompt}"

    # 构建消息（支持多模态）
    if _is_multimodal_llm(llm) and images:
        content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
        for img_path in images:
            if not os.path.exists(img_path):
                continue
            mime_type, _ = mimetypes.guess_type(img_path)
            if not mime_type:
                mime_type = "image/jpeg"
            base64_str = _encode_image(img_path)
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
        print(f"{'=' * 60}")
        print(f"📝 预处理讲稿请求")
        print(f"{'=' * 60}\n")

    response = llm.generate(messages, temperature=0.3)

    if DEBUG_LLM:
        print(f"\n{'=' * 60}")
        print("📥 [DEBUG] LLM 响应 (llm_preprocess_script)")
        print(f"{'=' * 60}")
        print(f"{response[:1000]}...")
        print(f"{'=' * 60}\n")

    # 清理响应：移除可能的 markdown 代码块标记
    result = response.strip()
    if result.startswith("```"):
        # 移除开头的 ```markdown 或 ```
        lines = result.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        result = "\n".join(lines)

    return result


def _parse_preprocessed_script(
    preprocessed_text: str,
    image_dir: Path,
) -> List[Dict]:
    """
    解析预处理后的带标记讲稿，返回 blocks 列表。

    Args:
        preprocessed_text: 带【PPT】标记的讲稿文本
        image_dir: 图片目录

    Returns:
        blocks 列表，每个 block 包含 template_hint, text, images
    """
    blocks: List[Dict] = []
    current_block: Optional[Dict] = None

    # 按行解析
    for line in preprocessed_text.split("\n"):
        marker_match = MARKER_RE.match(line.strip())
        if marker_match:
            # 保存之前的 block
            if current_block and current_block.get("text", "").strip():
                blocks.append(current_block)
            # 开始新的 block
            template_num = int(marker_match.group(1))
            current_block = {
                "template_hint": template_num,
                "text": "",
                "images": [],
            }
        elif current_block is not None:
            # 检查是否有图片引用
            img_match = re.search(r"\[图片资源:\s*([^\]]+)\]", line)
            if img_match:
                img_name = img_match.group(1).strip()
                img_path = image_dir / img_name
                if img_path.exists():
                    current_block["images"].append(str(img_path))
                # 从文本中移除图片标记
                line = re.sub(r"\[图片资源:\s*[^\]]+\]", "", line)

            current_block["text"] += line + "\n"

    # 保存最后一个 block
    if current_block and current_block.get("text", "").strip():
        blocks.append(current_block)

    # 清理每个 block 的文本
    for block in blocks:
        block["text"] = block["text"].strip()

    return blocks


def _extract_json_value(text: str, opener: str) -> Any:
    decoder = json.JSONDecoder()
    idx = 0
    while idx < len(text):
        start = text.find(opener, idx)
        if start == -1:
            break
        try:
            value, offset = decoder.raw_decode(text[start:])
            return value
        except json.JSONDecodeError:
            idx = start + 1
    raise ValueError("模型输出中未找到 JSON")


def _ensure_json_object(text: str) -> Dict:
    value = _extract_json_value(text.strip(), "{")
    if not isinstance(value, dict):
        raise ValueError("解析结果不是 JSON 对象")
    return value


def _ensure_json_array(text: str) -> List[Dict]:
    value = _extract_json_value(text.strip(), "[")
    if not isinstance(value, list):
        raise ValueError("解析结果不是 JSON 数组")
    return value


def _coerce_dict(entry):
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
                return _coerce_dict(candidate)
            except ValueError:
                continue
        raise ValueError("列表元素中未找到 JSON 对象。")
    raise ValueError("模型输出的元素不是有效的 JSON 对象。")


def choose_llm(
    enable: bool,
    provider: str,
    model: Optional[str],
    base_url: Optional[str] = None,
) -> Optional[BaseLLM]:
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
            raise ValueError(
                "Qwen provider 需要提供 --llm-base-url 或设置 QWEN_VLLM_BASE_URL。"
            )
        return QwenVLLM(base_url=endpoint)
    if provider == "taichu":
        final_model = model or "taichu4_vl_32b"
        return TaichuLLM(model=final_model, base_url=base_url)
    if provider == "glm" or provider == "zhipu":
        final_model = model or "glm-4.5v"
        return GLMLLM(model=final_model, base_url=base_url)
    raise ValueError(f"暂不支持的大模型提供商：{provider}")


def _apply_metadata_overrides(content: Dict, template_info: Dict, metadata: Dict):
    if not metadata:
        return
    for field in template_info["text_fields"]:
        path = list(field["path"])
        key = "/".join(path)
        value = None
        if metadata.get("lecturer") and "主讲" in key:
            value = metadata["lecturer"]
        elif metadata.get("college") and "学院" in key:
            value = metadata["college"]
        elif metadata.get("course") and any(
            token in key for token in ("课程", "课程名称", "项目")
        ):
            value = metadata["course"]
        if value is not None:
            _assign_in_schema(content, path, value)


def _fill_with_template(
    template_num: int,
    template_info: Dict,
    block: Dict,
    llm: Optional[BaseLLM],
    metadata: Dict,
    user_prompt: Optional[str] = None,
    use_multimodal: bool = True,
) -> Dict:
    """
    使用模板填充单个 block 的内容。

    Args:
        use_multimodal: 是否使用多模态消息（默认 True）
                       当讲稿有 PPT 标记时，图片位置已确定，建议设为 False
    """
    content = llm_fill_slide(
        llm,
        template_info,
        block.get("text", ""),
        block.get("images", []),
        user_prompt,
        use_multimodal,
    )
    _apply_metadata_overrides(content, template_info, metadata)
    return {
        "page_type": template_info["page_type"],
        "template_page_num": template_num,
        "content": content,
    }


def _strip_values(node):
    if isinstance(node, dict):
        if "type" in node and "value" in node:
            return node.get("value", "")
        return {k: _strip_values(v) for k, v in node.items()}
    if isinstance(node, list):
        return [_strip_values(item) for item in node]
    return node


def _empty_content(template_info: Dict) -> Dict:
    result = _clone_schema(template_info["schema"])
    for field in template_info["text_fields"]:
        _assign_in_schema(result, list(field["path"]), "")
    for field in template_info["image_fields"]:
        _assign_in_schema(result, list(field["path"]), "")
    return result


def _prepend_cover_page(pages: List[Dict], templates: Dict[int, Dict]):
    cover_template = templates.get(1)
    if not cover_template:
        return
    if pages and pages[0].get("template_page_num") == 1:
        return
    pages.insert(
        0,
        {
            "page_type": cover_template["page_type"],
            "template_page_num": 1,
            "content": _empty_content(cover_template),
        },
    )


def _fill_by_markers(
    blocks: List[Dict],
    templates: Dict[int, Dict],
    llm: Optional[BaseLLM],
    metadata: Dict,
    user_prompt: Optional[str] = None,
) -> List[Dict]:
    """
    按照讲稿中的 PPT 标记填充内容。

    由于讲稿已有明确的标记，图片位置已经确定（每个 block 的 images 字段），
    因此不需要使用多模态模型来界定图片位置，设置 use_multimodal=False。
    """
    pages: List[Dict] = []
    for block in blocks:
        template_num = block.get("template_hint")
        if template_num is None:
            continue
        if template_num not in templates:
            raise ValueError(
                f"模板 {template_num} 未在 template.json 中定义或不在 template.txt 中允许。"
            )
        pages.append(
            _fill_with_template(
                template_num,
                templates[template_num],
                block,
                llm,
                metadata,
                user_prompt,
                use_multimodal=False,  # 有标记时图片位置已确定，不需要多模态
            )
        )
    return pages


def _plan_without_markers(
    blocks: List[Dict],
    templates: Dict[int, Dict],
    llm: BaseLLM,
    metadata: Dict,
    user_prompt: Optional[str] = None,
    run_dir: Optional[Path] = None,
) -> List[Dict]:
    """
    处理没有【PPT】标记的讲稿。

    新流程（两步处理）：
    1. 预分页：让 LLM 将原始讲稿改写为带【PPT】标记的中间讲稿
    2. 填充：复用 _fill_by_markers 处理中间讲稿

    Args:
        blocks: 原始讲稿的 block 列表
        templates: 模板定义字典
        llm: LLM 实例
        metadata: 元数据
        user_prompt: 用户自定义提示
        run_dir: 运行目录，用于保存中间讲稿

    Returns:
        填充后的页面列表
    """
    if not llm:
        raise ValueError("讲稿未指定 PPT 标记且未启用 LLM，无法自动分配模板。")

    # 合并所有 block 的文本和图片
    doc_text = "\n\n".join(
        block.get("text", "") for block in blocks if block.get("text")
    )
    all_images = [path for block in blocks for path in block.get("images", [])]

    # Step 1: 预分页 - 生成带【PPT】标记的中间讲稿
    print("📝 Step 1: 预处理讲稿（生成带标记的中间讲稿）...")
    preprocessed_script = llm_preprocess_script(
        llm, doc_text, templates, all_images, user_prompt
    )

    # 保存中间讲稿到文件（供管理员/开发者下载）
    if run_dir:
        script_path = run_dir / "preprocessed_script.md"
        script_path.write_text(preprocessed_script, encoding="utf-8")
        print(f"💾 中间讲稿已保存: {script_path}")

    # Step 2: 解析中间讲稿为 blocks
    image_dir = run_dir / "images" if run_dir else Path(".")
    preprocessed_blocks = _parse_preprocessed_script(preprocessed_script, image_dir)

    if not preprocessed_blocks:
        raise ValueError("预处理后的讲稿没有有效的【PPT】标记，请检查 LLM 输出。")

    print(f"✅ 预处理完成，共 {len(preprocessed_blocks)} 个页面")

    # Step 3: 复用 _fill_by_markers 处理
    print("📝 Step 2: 填充页面内容...")
    pages = _fill_by_markers(preprocessed_blocks, templates, llm, metadata, user_prompt)

    return pages


def generate_config_data(
    docx_path: str,
    template_json: str,
    template_list: str,
    use_llm: bool,
    llm_provider: str,
    llm_model: Optional[str],
    llm_base_url: Optional[str],
    metadata_overrides: Optional[Dict[str, str]],
    run_dir: Path,
    user_prompt: Optional[str] = None,
) -> Dict:
    """核心逻辑：生成 JSON 内容，供 GUI/CLI 复用。"""
    metadata_overrides = metadata_overrides or {}
    image_dir = run_dir / "images"
    blocks, has_marker, metadata = parse_docx_blocks(docx_path, image_dir)
    for key in ("course", "college", "lecturer"):
        if metadata_overrides.get(key):
            metadata[key] = metadata_overrides[key]

    templates = load_template_defs(template_json, template_list)
    llm = choose_llm(use_llm, llm_provider, llm_model, llm_base_url)

    if has_marker:
        pages = _fill_by_markers(blocks, templates, llm, metadata, user_prompt)
    else:
        pages = _plan_without_markers(
            blocks, templates, llm, metadata, user_prompt, run_dir
        )

    if not pages:
        raise ValueError("未生成任何幻灯片内容，请检查讲稿或模板。")

    _prepend_cover_page(pages, templates)

    stripped_pages = []
    for page in pages:
        stripped_pages.append(
            {
                "page_type": page.get("page_type"),
                "template_page_num": page.get("template_page_num"),
                "content": _strip_values(page.get("content", {})),
            }
        )

    return {"ppt_pages": stripped_pages}


def process_docx(
    docx_path: str,
    template_json: str,
    template_list: str,
    output_path: Optional[str],
    use_llm: bool,
    llm_provider: str,
    llm_model: Optional[str],
    llm_base_url: Optional[str],
    override_course: Optional[str],
    override_college: Optional[str],
    override_lecturer: Optional[str],
    run_dir: Optional[str],
    config_name: str,
):
    """CLI 包装：处理参数、保证 run 目录存在，并额外复制文件到 output。"""
    metadata_overrides = {
        "course": override_course,
        "college": override_college,
        "lecturer": override_lecturer,
    }

    base_dir = Path(run_dir) if run_dir else _create_run_dir()
    base_dir.mkdir(parents=True, exist_ok=True)
    config_path = base_dir / config_name

    config = generate_config_data(
        docx_path,
        template_json,
        template_list,
        use_llm,
        llm_provider,
        llm_model,
        llm_base_url,
        metadata_overrides,
        base_dir,
    )

    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if output_path:
        explicit = Path(output_path)
        explicit.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(config_path, explicit)
        print(f"📄 另存为：{explicit}")

    print(f"✅ 已生成 JSON：{config_path}")
    print(f"📁 资源输出目录：{base_dir}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="根据 DOCX 讲稿生成 PPT 配置 JSON。")
    parser.add_argument("--docx", required=True, help="讲稿 DOCX 路径")
    parser.add_argument(
        "--template-json", default="template/template.json", help="模板定义 JSON 文件"
    )
    parser.add_argument(
        "--template-list", default="template/template.txt", help="可用模板编号列表 txt"
    )
    parser.add_argument(
        "--output",
        default=None,
        help="如需额外复制一份 JSON，请提供完整路径；若省略则仅在 temp/run-*/ 中生成",
    )
    parser.add_argument("--use-llm", action="store_true", help="启用大模型填充/排版")
    parser.add_argument("--llm-provider", default="deepseek", help="大模型提供商")
    parser.add_argument("--llm-model", default="deepseek-chat", help="大模型名称")
    parser.add_argument(
        "--llm-base-url",
        default="http://172.18.75.58:9000",
        help="自定义大模型接口地址",
    )
    parser.add_argument("--course-name", default=None, help="手动指定课程/项目名称")
    parser.add_argument("--college-name", default=None, help="手动指定学院/单位")
    parser.add_argument("--lecturer-name", default=None, help="手动指定主讲教师姓名")
    parser.add_argument(
        "--run-dir",
        default=None,
        help="指定输出目录（默认在 temp 下自动创建 run-时间戳-随机值 文件夹）",
    )
    parser.add_argument(
        "--config-name",
        default="config.json",
        help="输出目录中生成的配置文件名称",
    )
    return parser


def main():
    args = build_arg_parser().parse_args()
    process_docx(
        docx_path=args.docx,
        template_json=args.template_json,
        template_list=args.template_list,
        output_path=args.output,
        use_llm=args.use_llm,
        llm_provider=args.llm_provider,
        llm_model=args.llm_model,
        llm_base_url=args.llm_base_url,
        override_course=args.course_name,
        override_college=args.college_name,
        override_lecturer=args.lecturer_name,
        run_dir=args.run_dir,
        config_name=args.config_name,
    )


def _strip_values(node):
    if isinstance(node, dict):
        if "type" in node and "value" in node:
            return node.get("value", "")
        return {k: _strip_values(v) for k, v in node.items()}
    if isinstance(node, list):
        return [_strip_values(item) for item in node]
    return node


if __name__ == "__main__":
    main()
