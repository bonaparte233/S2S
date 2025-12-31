"""DOCX 表格解析模块 - 将表格内容转换为结构化文本。

该模块提供通用的表格解析能力，不针对特定文档格式硬编码。
核心策略：
1. 智能去重合并单元格
2. 自动识别行类型（键值对、跨行内容、数据表等）
3. 输出为 Markdown 格式，便于 LLM 理解
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from docx.table import Table


@dataclass
class TableSection:
    """表格中的一个逻辑区块"""
    title: str = ""
    content: str = ""
    section_type: str = "text"  # text, key_value, data_table


@dataclass
class TableParseResult:
    """表格解析结果"""
    sections: List[TableSection] = field(default_factory=list)
    raw_text: str = ""  # 原始文本形式（兜底）


def _dedupe_row_cells(row) -> List[str]:
    """去重行中的单元格（处理合并单元格导致的重复）。
    
    合并单元格在 python-docx 中会返回相同的文本多次，
    通过去重可以得到实际的唯一内容。
    """
    unique_cells = []
    seen_texts = set()
    for cell in row.cells:
        text = cell.text.strip()
        # 使用 (text, cell位置) 的方式避免误删相同内容的不同单元格
        # 但对于合并单元格，连续相同内容应该去重
        if text not in seen_texts or text == "":
            unique_cells.append(text)
            if text:  # 空字符串不加入 seen，允许多个空单元格
                seen_texts.add(text)
    return unique_cells


def _is_header_row(cells: List[str]) -> bool:
    """判断是否为表头行。
    
    通用规则：包含常见表头关键词
    """
    header_keywords = {"序号", "编号", "名称", "描述", "状态", "类型", "项目", "内容"}
    if not cells:
        return False
    # 如果第一个单元格是"序号"或类似，很可能是表头
    first = cells[0].lower()
    if first in {"序号", "编号", "no", "no.", "#"}:
        return True
    # 如果多个单元格都是短文本（可能是列标题）
    short_cells = [c for c in cells if c and len(c) <= 10]
    if len(short_cells) >= 2 and len(short_cells) == len([c for c in cells if c]):
        return any(kw in "".join(cells) for kw in header_keywords)
    return False


def _is_data_row(cells: List[str]) -> bool:
    """判断是否为数据行（首列为数字序号）。"""
    if not cells:
        return False
    first = cells[0].strip()
    return first.isdigit() or (len(first) <= 3 and first.replace(".", "").isdigit())


def _extract_title_from_content(text: str) -> Tuple[str, str]:
    """从跨行内容中提取标题和正文。
    
    支持格式：
    - "标题：内容"
    - "标题(说明)：内容"
    - "标题\n内容"
    """
    if not text:
        return "", ""
    
    # 尝试按冒号分割
    for sep in ["：", ":"]:
        if sep in text:
            parts = text.split(sep, 1)
            title_part = parts[0].strip()
            content_part = parts[1].strip() if len(parts) > 1 else ""
            # 标题不应该太长
            if len(title_part) <= 30:
                # 清理标题中的括号说明
                clean_title = title_part.split("(")[0].split("（")[0].strip()
                return clean_title, content_part
    
    # 尝试按换行分割
    lines = text.split("\n", 1)
    if len(lines) > 1 and len(lines[0]) <= 30:
        return lines[0].strip(), lines[1].strip()
    
    return "", text


def _identify_row_type(cells: List[str]) -> str:
    """识别行类型。
    
    Returns:
        'key_value': 键值对（2个单元格）
        'multi_kv': 多键值对（4个单元格，交替键值）
        'span': 跨行内容（1个单元格）
        'header': 表头行
        'data': 数据行
        'matrix': 矩阵数据行（如分工表）
        'unknown': 未知类型
    """
    if not cells:
        return "unknown"
    
    non_empty = [c for c in cells if c.strip()]
    count = len(non_empty)
    
    if count == 0:
        return "unknown"
    elif count == 1:
        return "span"
    elif count == 2:
        return "key_value"
    elif count == 4:
        # 检查是否为交替键值对（键通常较短）
        if all(len(cells[i]) <= 15 for i in [0, 2] if i < len(cells) and cells[i]):
            return "multi_kv"
    
    if _is_header_row(cells):
        return "header"
    if _is_data_row(cells):
        return "data"
    
    # 检查是否为百分比矩阵行
    pct_count = sum(1 for c in cells if "%" in c)
    if pct_count >= 2:
        return "matrix"
    
    return "unknown"


def _format_key_value(key: str, value: str) -> str:
    """格式化键值对为 Markdown。"""
    key = key.strip()
    value = value.strip()
    if not key and not value:
        return ""
    if not value:
        return f"**{key}**"
    return f"**{key}**：{value}"


def _format_data_table(headers: List[str], rows: List[List[str]]) -> str:
    """格式化数据表为 Markdown 表格。"""
    if not headers and not rows:
        return ""
    
    # 如果没有表头，用第一行作为表头
    if not headers and rows:
        headers = rows[0]
        rows = rows[1:]
    
    if not headers:
        return ""
    
    lines = []
    # 表头
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    # 数据行
    for row in rows:
        # 补齐列数
        padded = row + [""] * (len(headers) - len(row))
        lines.append("| " + " | ".join(padded[:len(headers)]) + " |")
    
    return "\n".join(lines)


def parse_single_table(table: Table) -> TableParseResult:
    """解析单个表格，返回结构化结果。
    
    Args:
        table: python-docx 的 Table 对象
        
    Returns:
        TableParseResult 包含解析后的区块列表
    """
    result = TableParseResult()
    sections: List[TableSection] = []
    
    # 用于收集数据表的临时变量
    current_headers: List[str] = []
    current_data_rows: List[List[str]] = []
    
    def flush_data_table():
        """将累积的数据表转为 section"""
        nonlocal current_headers, current_data_rows
        if current_headers or current_data_rows:
            content = _format_data_table(current_headers, current_data_rows)
            if content:
                sections.append(TableSection(
                    title="数据表",
                    content=content,
                    section_type="data_table"
                ))
        current_headers = []
        current_data_rows = []
    
    for row in table.rows:
        cells = _dedupe_row_cells(row)
        row_type = _identify_row_type(cells)
        
        if row_type == "span":
            flush_data_table()
            text = cells[0] if cells else ""
            title, content = _extract_title_from_content(text)
            sections.append(TableSection(
                title=title,
                content=content or text,
                section_type="text"
            ))
            
        elif row_type == "key_value":
            flush_data_table()
            content = _format_key_value(cells[0], cells[1])
            sections.append(TableSection(
                title=cells[0].strip(),
                content=cells[1].strip(),
                section_type="key_value"
            ))
            
        elif row_type == "multi_kv":
            flush_data_table()
            # 交替键值对：[k1, v1, k2, v2]
            for i in range(0, len(cells) - 1, 2):
                key = cells[i].strip()
                value = cells[i + 1].strip() if i + 1 < len(cells) else ""
                if key or value:
                    sections.append(TableSection(
                        title=key,
                        content=value,
                        section_type="key_value"
                    ))
                    
        elif row_type == "header":
            flush_data_table()
            current_headers = [c.strip() for c in cells]
            
        elif row_type == "data":
            current_data_rows.append([c.strip() for c in cells])
            
        elif row_type == "matrix":
            # 矩阵行也当作数据行处理
            current_data_rows.append([c.strip() for c in cells])
            
        else:  # unknown
            # 尝试作为数据行处理
            if current_headers:
                current_data_rows.append([c.strip() for c in cells])
            else:
                flush_data_table()
                # 作为普通文本
                text = " | ".join(c for c in cells if c.strip())
                if text:
                    sections.append(TableSection(
                        title="",
                        content=text,
                        section_type="text"
                    ))
    
    # 处理剩余的数据表
    flush_data_table()
    
    result.sections = sections
    result.raw_text = table_result_to_text(result)
    return result


def table_result_to_text(result: TableParseResult) -> str:
    """将解析结果转换为 Markdown 文本。
    
    Args:
        result: TableParseResult 对象
        
    Returns:
        Markdown 格式的文本
    """
    lines: List[str] = []
    
    for section in result.sections:
        if section.section_type == "key_value":
            lines.append(_format_key_value(section.title, section.content))
        elif section.section_type == "data_table":
            lines.append(section.content)
        elif section.section_type == "text":
            if section.title:
                lines.append(f"### {section.title}")
            if section.content:
                lines.append(section.content)
        lines.append("")  # 空行分隔
    
    return "\n".join(lines).strip()


def extract_tables_from_docx(doc) -> List[TableParseResult]:
    """从 DOCX 文档中提取所有表格。
    
    Args:
        doc: python-docx 的 Document 对象
        
    Returns:
        TableParseResult 列表
    """
    results: List[TableParseResult] = []
    
    for table in doc.tables:
        try:
            result = parse_single_table(table)
            if result.sections:  # 只添加非空结果
                results.append(result)
        except Exception as e:
            # 解析失败时记录警告，继续处理其他表格
            import logging
            logging.getLogger(__name__).warning(f"表格解析失败: {e}")
            continue
    
    return results
