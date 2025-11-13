"""
Word 文档解析器 (S1)

功能: 解析 .docx 讲稿，提取文本、图片，并将列表/表格转为 Markdown
职责: 只负责 Word 文档解析，不涉及 LLM 或 PPT 生成 (SRP 原则)
输出: raw_data.json
"""

import re
import json
import os
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


# 图片存储目录 (配置驱动)
IMG_DIR = "temp_data/extracted_images"


def table_to_markdown(table: Table) -> str:
    """
    将 Word 表格转换为 Markdown 表格格式

    :param table: Word 表格对象
    :return: Markdown 格式的表格字符串
    """
    markdown_lines = []

    for row_idx, row in enumerate(table.rows):
        # 提取单元格文本
        cells_text = [cell.text.strip() for cell in row.cells]
        markdown_lines.append("| " + " | ".join(cells_text) + " |")

        # 在表头后添加分隔线
        if row_idx == 0:
            markdown_lines.append("|" + "|".join(["---"] * len(cells_text)) + "|")

    return "\n".join(markdown_lines)


def get_list_level_markdown(para: Paragraph) -> str:
    """
    根据段落样式和缩进，生成 Markdown 列表前缀

    :param para: Word 段落对象
    :return: Markdown 列表前缀 (如 "- ", "  - ", "    - ")
    """
    # 检查是否为列表样式
    if para.style.name and "List" in para.style.name:
        # 获取缩进级别 (假设每级缩进为 2 个空格)
        indent = para.paragraph_format.left_indent
        if indent:
            # 将缩进转换为级别 (每 457200 EMU 约为一级缩进)
            level = int(indent / 457200)
            return "  " * level + "- "
        else:
            return "- "

    return ""


def parse_word(doc_path: str, output_path: str = "temp_data/raw_data.json") -> str:
    """
    解析 Word 文档，提取文本、图片和表格

    :param doc_path: Word 文档路径
    :param output_path: 输出的 raw_data.json 路径
    :return: 输出文件路径
    """
    # 输入验证 (防御性编程)
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Word 文档不存在: {doc_path}")

    # 创建图片存储目录
    os.makedirs(IMG_DIR, exist_ok=True)

    # 初始化状态变量
    slides_data = []
    current_layout_id = None
    current_text_buffer = []
    current_images_buffer = []
    image_counter = 0

    # 加载 Word 文档
    doc = Document(doc_path)

    # 布局标记正则表达式
    layout_regex = re.compile(r"【PPT(\d+)】")

    def flush_slide():
        """
        内部函数: 保存当前幻灯片数据并重置缓冲区
        """
        nonlocal current_layout_id, current_text_buffer, current_images_buffer

        if current_layout_id:
            raw_text = "\n".join(current_text_buffer)
            slides_data.append(
                {
                    "layout_id": current_layout_id,
                    "raw_text": raw_text,
                    "raw_images": list(current_images_buffer),
                }
            )
            current_text_buffer.clear()
            current_images_buffer.clear()

    # 遍历文档的所有元素 (段落和表格)
    for element in doc.element.body:
        # 处理段落
        if element.tag.endswith("p"):
            para = Paragraph(element, doc)

            # 检查是否为布局标记
            match = layout_regex.search(para.text)

            if match:
                # 遇到新标记，保存上一页
                flush_slide()

                # 重置当前布局 ID
                current_layout_id = match.group(1)

                # 提取标记后的文本
                text_after_marker = para.text[match.end() :].strip()
                if text_after_marker:
                    current_text_buffer.append(text_after_marker)
            else:
                # 普通文本段落
                if current_layout_id:  # 只在已开始某个幻灯片后才添加
                    # 处理列表格式
                    list_prefix = get_list_level_markdown(para)
                    current_text_buffer.append(f"{list_prefix}{para.text}")

            # 检查段落中的图片 (无论是否为标记段落)
            if current_layout_id:
                for run in para.runs:
                    # 检查 run 中是否包含图片
                    for drawing in run._element.findall(
                        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"
                    ):
                        for blip in drawing.findall(
                            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
                        ):
                            # 获取图片的 rId
                            rId = blip.get(
                                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                            )
                            if rId:
                                # 从文档关系中获取图片
                                image_part = doc.part.related_parts[rId]
                                image_blob = image_part.blob

                                # 保存图片
                                img_path = f"{IMG_DIR}/img_{image_counter}.png"
                                with open(img_path, "wb") as f:
                                    f.write(image_blob)

                                # 记录图片路径
                                current_images_buffer.append(img_path)
                                current_text_buffer.append(f"[IMAGE: {img_path}]")
                                image_counter += 1

        # 处理表格
        elif element.tag.endswith("tbl"):
            if current_layout_id:  # 只在已开始某个幻灯片后才添加
                table = Table(element, doc)
                md_table = table_to_markdown(table)
                current_text_buffer.append(md_table)

    # 收尾: 保存最后一页
    flush_slide()

    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    # 输出 JSON 文件
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(slides_data, f, indent=2, ensure_ascii=False)

    print(f"✓ Word 解析完成，共提取 {len(slides_data)} 页幻灯片")
    print(f"✓ 输出文件: {output_path}")

    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="解析 Word 讲稿并提取内容")
    parser.add_argument("--doc", required=True, help="Word 文档路径")
    parser.add_argument(
        "--output", default="temp_data/raw_data.json", help="输出 JSON 文件路径"
    )

    args = parser.parse_args()

    try:
        parse_word(args.doc, args.output)
    except Exception as e:
        print(f"✗ 错误: {e}")
        exit(1)
