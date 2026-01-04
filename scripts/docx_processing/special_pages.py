"""特殊页面处理模块（封面、目录、结束页）。"""

from __future__ import annotations

from typing import Dict, List, Optional

from scripts.llm_client import BaseLLM
from scripts.docx_processing.template_utils import (
    assign_in_schema,
    clone_schema,
    empty_content,
)
from scripts.docx_processing.llm_processor import llm_fill_slide


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


def prepend_cover_page(
    pages: List[Dict],
    templates: Dict[int, Dict],
    metadata: Optional[Dict] = None,
    llm: Optional[BaseLLM] = None,
    global_config: Optional[Dict] = None,
) -> None:
    """在页面列表开头插入封面页。"""
    # 从配置读取封面页码
    special_pages = (global_config or {}).get("special_pages", {})
    cover_page_num = special_pages.get("cover")

    # 如果未配置封面页，跳过
    if cover_page_num is None:
        print("ℹ️  未配置封面页，跳过封面页生成")
        return

    cover_template = templates.get(cover_page_num)
    if not cover_template:
        print(f"⚠️  模板{cover_page_num}（封面页）不存在，跳过封面页生成")
        return

    if pages and pages[0].get("template_page_num") == cover_page_num:
        if metadata:
            _apply_metadata_overrides(
                pages[0].get("content", {}), cover_template, metadata
            )
        return

    if llm and metadata:
        cover_text = "这是封面页，请根据预设元信息填充相应字段。"
        content = llm_fill_slide(
            llm, cover_template, cover_text, [], None, False, metadata
        )
    else:
        content = empty_content(cover_template)
        if metadata:
            _apply_metadata_overrides(content, cover_template, metadata)

    pages.insert(0, {
        "page_type": cover_template["page_type"],
        "template_page_num": cover_page_num,
        "content": content,
    })


def insert_toc_page(
    pages: List[Dict],
    templates: Dict[int, Dict],
    metadata: Dict,
    chapters: List[str],
    global_config: Optional[Dict] = None,
) -> None:
    """在封面页后插入目录页。"""
    # 从配置读取目录页码和封面页码
    special_pages = (global_config or {}).get("special_pages", {})
    toc_page_num = special_pages.get("toc")
    cover_page_num = special_pages.get("cover")

    if toc_page_num is None:
        print("ℹ️  未配置目录页，跳过目录页生成")
        return

    toc_template = templates.get(toc_page_num)
    if not toc_template:
        print(f"⚠️  模板{toc_page_num}（目录页）不存在，跳过目录页生成")
        return

    if not chapters:
        print("ℹ️  未检测到章节信息，跳过目录页生成")
        return

    content = clone_schema(toc_template["schema"])

    # 填充总课程名称
    course_name = metadata.get("course", "")
    for field in toc_template["text_fields"]:
        path = list(field["path"])
        key = "/".join(path)
        if "总课程名称" in key or "课程" in key:
            assign_in_schema(content, path, course_name)
            break

    # 填充章节标题（最多4个）
    for idx, chapter in enumerate(chapters[:4], start=1):
        for field in toc_template["text_fields"]:
            path = list(field["path"])
            key = "/".join(path)
            if f"章节标题{idx}" in key:
                assign_in_schema(content, path, chapter)
                break
        for field in toc_template["text_fields"]:
            path = list(field["path"])
            key = "/".join(path)
            if f"章节内容{idx}" in key:
                assign_in_schema(content, path, "")
                break

    # 清空未使用的章节字段
    for idx in range(len(chapters) + 1, 5):
        for field in toc_template["text_fields"]:
            path = list(field["path"])
            key = "/".join(path)
            if f"章节标题{idx}" in key or f"章节内容{idx}" in key:
                assign_in_schema(content, path, "")

    toc_page = {
        "page_type": toc_template["page_type"],
        "template_page_num": toc_page_num,
        "content": content,
    }

    # 如果有封面页且第一页是封面页，则插入到封面页后面
    insert_pos = 0
    if cover_page_num is not None and pages and pages[0].get("template_page_num") == cover_page_num:
        insert_pos = 1
    pages.insert(insert_pos, toc_page)
    print(f"✅ 已插入目录页，包含 {len(chapters[:4])} 个章节")


def append_end_page(
    pages: List[Dict],
    templates: Dict[int, Dict],
    metadata: Dict,
    global_config: Optional[Dict] = None,
) -> None:
    """在页面列表末尾添加结束页。"""
    # 从配置读取结束页码
    special_pages = (global_config or {}).get("special_pages", {})
    end_page_num = special_pages.get("end")

    if end_page_num is None:
        print("ℹ️  未配置结束页，跳过结束页生成")
        return

    end_template = templates.get(end_page_num)
    if not end_template:
        print(f"⚠️  模板{end_page_num}（结束页）不存在，跳过结束页生成")
        return

    content = clone_schema(end_template["schema"])

    course_name = metadata.get("course", "")
    for field in end_template["text_fields"]:
        path = list(field["path"])
        key = "/".join(path)
        if "总课程名称" in key or "课程" in key:
            assign_in_schema(content, path, course_name)
            break

    end_page = {
        "page_type": end_template["page_type"],
        "template_page_num": end_page_num,
        "content": content,
    }

    pages.append(end_page)
    print("✅ 已添加结束页")
