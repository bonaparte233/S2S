"""幻灯片内容填充模块。"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from scripts.llm_client import BaseLLM
from scripts.docx_processing.constants import MARKER_RE
from scripts.docx_processing.docx_parser import SectionContext
from scripts.docx_processing.template_utils import (
    assign_in_schema,
    get_in_schema,
)
from scripts.docx_processing.llm_processor import (
    llm_fill_slide,
    llm_preprocess_script,
)


def _apply_metadata_overrides(
    content: Dict, template_info: Dict, metadata: Dict
) -> None:
    """应用元数据覆盖。"""
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
            assign_in_schema(content, path, value)


def _apply_section_context(
    content: Dict,
    template_info: Dict,
    section_context: Dict[str, Optional[str]],
    title_hint: Optional[str] = None,
    global_config: Optional[Dict] = None,
) -> None:
    """根据章节上下文填充章节相关字段。"""
    if not section_context:
        return

    # 从全局配置读取 section_tracking
    template_prompt = (global_config or {}).get("template_prompt", {})
    if not template_prompt.get("section_tracking", True):
        return

    # 默认字段映射（只有两级：章节和知识点）
    field_mappings = [
        (["章节", "一级标题", "章节名", "大章节"], "chapter"),
        (["知识点", "二级标题", "小节名", "节名"], "section"),
    ]

    def _trim_to_max(value: str, max_chars: Optional[int]) -> str:
        if max_chars and len(value) > max_chars:
            return value[:max_chars]
        return value

    def _norm(text: Optional[str]) -> str:
        return re.sub(r"[（()）\s]", "", text or "")

    for field in template_info["text_fields"]:
        path = list(field["path"])
        key = "/".join(path)
        max_chars = field.get("max_chars")
        current_value = get_in_schema(content, path) or ""
        required = bool(field.get("required"))

        for keywords, ctx_key in field_mappings:
            if any(kw in key for kw in keywords):
                target = section_context.get(ctx_key)
                if not target:
                    break
                target = _trim_to_max(str(target), max_chars)
                if required:
                    if _norm(current_value) != _norm(target):
                        assign_in_schema(content, path, target)
                else:
                    if not current_value:
                        assign_in_schema(content, path, target)
                break


def _fill_with_template(
    template_num: int,
    template_info: Dict,
    block: Dict,
    llm: Optional[BaseLLM],
    metadata: Dict,
    user_prompt: Optional[str] = None,
    use_multimodal: bool = True,
    section_context: Optional[Dict[str, Optional[str]]] = None,
    global_config: Optional[Dict] = None,
) -> Dict:
    """使用模板填充单个 block 的内容。"""
    content = llm_fill_slide(
        llm,
        template_info,
        block.get("text", ""),
        block.get("images", []),
        user_prompt,
        use_multimodal,
        metadata,
        section_context,
        global_config,
    )
    _apply_metadata_overrides(content, template_info, metadata)
    return {
        "page_type": template_info["page_type"],
        "template_page_num": template_num,
        "content": content,
    }


def _parse_preprocessed_script(
    preprocessed_text: str, image_dir: Path
) -> List[Dict]:
    """解析预处理后的带标记讲稿。"""
    blocks: List[Dict] = []
    current_block: Optional[Dict] = None

    for line in preprocessed_text.split("\n"):
        marker_match = MARKER_RE.match(line.strip())
        if marker_match:
            if current_block and current_block.get("text", "").strip():
                blocks.append(current_block)
            template_num = int(marker_match.group(1))
            current_block = {
                "template_hint": template_num,
                "text": "",
                "images": [],
            }
        elif current_block is not None:
            img_match = re.search(r"\[图片资源:\s*([^\]]+)\]", line)
            if img_match:
                img_ref = img_match.group(1).strip()
                if os.path.isabs(img_ref) and Path(img_ref).exists():
                    current_block["images"].append(img_ref)
                else:
                    img_name = Path(img_ref).name
                    img_path = image_dir / img_name
                    if img_path.exists():
                        current_block["images"].append(str(img_path))
                line = re.sub(r"\[图片资源:\s*[^\]]+\]", "", line)
            current_block["text"] += line + "\n"

    if current_block and current_block.get("text", "").strip():
        blocks.append(current_block)

    for block in blocks:
        block["text"] = block["text"].strip()

    return blocks


def fill_by_markers(
    blocks: List[Dict],
    templates: Dict[int, Dict],
    llm: Optional[BaseLLM],
    metadata: Dict,
    user_prompt: Optional[str] = None,
    global_config: Optional[Dict] = None,
) -> Tuple[List[Dict], List[str]]:
    """按照讲稿中的 PPT 标记填充内容。"""
    pages: List[Dict] = []
    section_ctx = SectionContext()
    chapters: List[str] = []

    # 从全局配置读取 section_tracking
    template_prompt = (global_config or {}).get("template_prompt", {})
    section_tracking = template_prompt.get("section_tracking", True)

    for block in blocks:
        template_num = block.get("template_hint")
        if template_num is None:
            continue
        if template_num not in templates:
            raise ValueError(f"模板 {template_num} 未在 template.json 中定义。")

        template_info = templates[template_num]
        page_type = template_info.get("page_type", "")
        block_text = block.get("text", "")
        lines = [l.strip() for l in block_text.split("\n") if l.strip()]

        # 仅在启用 section_tracking 时更新章节上下文（只追踪两级）
        if section_tracking:
            if "章节" in page_type and lines:
                section_ctx.level1 = lines[0]
                section_ctx.level2 = None
                if len(lines) > 1:
                    section_ctx.level2 = lines[1]
                if section_ctx.level1 and section_ctx.level1 not in chapters:
                    chapters.append(section_ctx.level1)
            else:
                for line in lines[:5]:
                    old_level1 = section_ctx.level1
                    section_ctx.update(line)
                    if section_ctx.level1 and section_ctx.level1 != old_level1:
                        if section_ctx.level1 not in chapters:
                            chapters.append(section_ctx.level1)

        pages.append(
            _fill_with_template(
                template_num,
                template_info,
                block,
                llm,
                metadata,
                user_prompt,
                use_multimodal=False,
                section_context=section_ctx.to_dict() if section_tracking else None,
                global_config=global_config,
            )
        )
    return pages, chapters


def preprocess_and_fill(
    blocks: List[Dict],
    templates: Dict[int, Dict],
    llm: Optional[BaseLLM],
    metadata: Dict,
    user_prompt: Optional[str] = None,
    run_dir: Optional[Path] = None,
    has_marker: bool = False,
    global_config: Optional[Dict] = None,
) -> Tuple[List[Dict], List[str]]:
    """统一的讲稿处理流程（两步处理）。"""
    if not llm:
        raise ValueError("未启用 LLM，无法处理讲稿。")

    # 合并所有 block 的文本和图片
    if has_marker:
        doc_parts = []
        for block in blocks:
            text = block.get("text", "")
            template_hint = block.get("template_hint")
            page_content = text
            for img_path in block.get("images", []):
                page_content += f"\n[图片资源: {img_path}]"
            if template_hint is not None:
                doc_parts.append(f"【PPT{template_hint}】\n{page_content}")
            elif page_content.strip():
                doc_parts.append(page_content)
        doc_text = "\n\n".join(doc_parts)
    else:
        doc_parts = []
        for block in blocks:
            text = block.get("text", "")
            for img_path in block.get("images", []):
                text += f"\n[图片资源: {img_path}]"
            if text.strip():
                doc_parts.append(text)
        doc_text = "\n\n".join(doc_parts)

    all_images = [path for block in blocks for path in block.get("images", [])]

    # Step 1: 预处理讲稿
    if has_marker:
        print("📝 Step 1: 优化讲稿文本（保持原有分页）...")
    else:
        print("📝 Step 1: 预处理讲稿（自动分页）...")

    preprocessed_script = llm_preprocess_script(
        llm, doc_text, templates, all_images, user_prompt, has_marker, global_config
    )

    # 保存中间讲稿
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

    # Step 3: 填充内容
    print("📝 Step 2: 填充页面内容...")
    pages, chapters = fill_by_markers(
        preprocessed_blocks, templates, llm, metadata, user_prompt, global_config
    )

    return pages, chapters
