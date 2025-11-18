"""根据 DOCX 讲稿和模板定义生成 JSON，可选调用 DeepSeek LLM。"""

from __future__ import annotations

import argparse
import json
import re
import secrets
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn

from llm_client import BaseLLM, DeepSeekLLM, LocalLLM

MARKER_RE = re.compile(r"【PPT(\d+)】")
IMAGE_NAME_TEMPLATE = "doc_image_{idx}.{ext}"


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


def _build_prompt(template_info: Dict, raw_text: str, images: List[str]) -> str:
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
    image_section = "无" if not images else "\n".join(images)
    meta = template_info.get("meta") or {}
    scene = "、".join(meta.get("scene", [])) or "通用"
    layout = meta.get("layout", template_info["page_type"])
    style = meta.get("style", "")
    note = meta.get("notes", "")
    prompt = f"""
请阅读以下讲稿并生成一个 JSON，对模板《{template_info["page_type"]}》的文本/图片字段进行填充。
模板布局：{layout}；使用场景：{scene}；风格提示：{style}
注意事项：{note}
务必记住讲稿中提到的主讲人姓名、课程/讲座/项目名称或其他关键专有名词，并在后续所有需要这些信息的字段保持完全一致、不要改写。所有标记为“required”的字段必须填写，且文本长度不得超过对应的 max_chars 限制。
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
    return prompt


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
    llm: BaseLLM, template_info: Dict, raw_text: str, images: List[str]
) -> Dict:
    if not llm:
        return _simple_fill(template_info, raw_text, images)

    prompt = _build_prompt(template_info, raw_text, images)
    response = llm.generate([{"role": "user", "content": prompt}], temperature=0.2)
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
    llm: BaseLLM, doc_text: str, templates: Dict[int, Dict], images: List[str]
) -> List[Dict]:
    template_desc = "\n".join(
        f"- 模板 {info['page_type']} (编号 {num}): 文本{len(info['text_fields'])}项, 图片{len(info['image_fields'])}项"
        for num, info in templates.items()
    )
    image_section = "无" if not images else "\n".join(images)
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

生成内容时务必记住并重复使用讲稿中的主讲人姓名、课程/讲座/项目名称等关键专有名词，确保在所有幻灯片中需要填写专有名词的位置保持一致，不要随意改写或另造新名称。所有标记为“required”的字段都必须提供文字，且不得超过对应的 max_chars 限制。
讲稿全文：
{doc_text}
"""
    response = llm.generate([{"role": "user", "content": prompt}], temperature=0.3)
    try:
        plan = _ensure_json_array(response)
        return plan
    except Exception:
        raise ValueError("模型输出无法解析为 JSON 数组，请检查提示或重试。")


def _ensure_json_object(text: str) -> Dict:
    text = text.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("模型输出中未找到 JSON 对象")
    return json.loads(text[start : end + 1])


def _ensure_json_array(text: str) -> List[Dict]:
    text = text.strip()
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1:
        raise ValueError("模型输出中未找到 JSON 数组")
    return json.loads(text[start : end + 1])


def choose_llm(enable: bool, provider: str, model: Optional[str]) -> Optional[BaseLLM]:
    if not enable:
        return None
    provider = (provider or "").lower()
    if provider == "deepseek":
        return DeepSeekLLM(model=model or "deepseek-chat")
    if provider == "local":
        return LocalLLM(model=model)
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
) -> Dict:
    content = llm_fill_slide(
        llm, template_info, block.get("text", ""), block.get("images", [])
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
) -> List[Dict]:
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
                template_num, templates[template_num], block, llm, metadata
            )
        )
    return pages


def _plan_without_markers(
    blocks: List[Dict], templates: Dict[int, Dict], llm: BaseLLM, metadata: Dict
) -> List[Dict]:
    if not llm:
        raise ValueError("讲稿未指定 PPT 标记且未启用 LLM，无法自动分配模板。")
    doc_text = "\n\n".join(
        block.get("text", "") for block in blocks if block.get("text")
    )
    all_images = [path for block in blocks for path in block.get("images", [])]
    plan = llm_plan_slides(llm, doc_text, templates, all_images)
    pages: List[Dict] = []
    for item in plan:
        template_num = item.get("template_page_num")
        if template_num not in templates:
            raise ValueError(f"模型返回的模板编号 {template_num} 不在允许列表中。")
        template_info = templates[template_num]
        content = _clone_schema(template_info["schema"])
        texts = item.get("texts", [])
        images = item.get("images", [])
        for idx, field in enumerate(template_info["text_fields"]):
            value = _lookup_field_value(
                field, texts, texts if isinstance(texts, list) else None, idx
            )
            _assign_in_schema(content, list(field["path"]), value)
        for idx, field in enumerate(template_info["image_fields"]):
            value = _lookup_field_value(
                field, images, images if isinstance(images, list) else None, idx
            )
            _assign_in_schema(content, list(field["path"]), value)
        _apply_metadata_overrides(content, template_info, metadata)
        pages.append(
            {
                "page_type": template_info["page_type"],
                "template_page_num": template_num,
                "content": content,
            }
        )
    return pages


def generate_config_data(
    docx_path: str,
    template_json: str,
    template_list: str,
    use_llm: bool,
    llm_provider: str,
    llm_model: Optional[str],
    metadata_overrides: Optional[Dict[str, str]],
    run_dir: Path,
) -> Dict:
    """核心逻辑：生成 JSON 内容，供 GUI/CLI 复用。"""
    metadata_overrides = metadata_overrides or {}
    image_dir = run_dir / "images"
    blocks, has_marker, metadata = parse_docx_blocks(docx_path, image_dir)
    for key in ("course", "college", "lecturer"):
        if metadata_overrides.get(key):
            metadata[key] = metadata_overrides[key]

    templates = load_template_defs(template_json, template_list)
    llm = choose_llm(use_llm, llm_provider, llm_model)

    if has_marker:
        pages = _fill_by_markers(blocks, templates, llm, metadata)
    else:
        pages = _plan_without_markers(blocks, templates, llm, metadata)

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
        metadata_overrides,
        base_dir,
    )

    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

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
