"""
幻灯片复制器 (使用 ZIP/XML 操作)

功能: 在 PPTX 包层面复制幻灯片，避免使用 python-pptx API 导致的重复元素问题
借鉴: ai2ppt 的 generatePPT_template.py 实现
职责: 只负责复制幻灯片，不涉及内容填充 (SRP 原则)
"""

import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# --- 全局常量：XML 命名空间 ---
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 注册命名空间，确保生成的 XML 格式正确
ET.register_namespace("", P_NS)
ET.register_namespace("r", OD_REL_NS)

# 正则表达式：匹配幻灯片和标签文件
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml")
TAG_RE = re.compile(r"ppt/tags/tag(\d+)\.xml")


def _next_rid(existing_ids):
    """生成下一个未被占用的 rId 序号"""
    nums = [
        int(rid[3:])
        for rid in existing_ids
        if rid.startswith("rId") and rid[3:].isdigit()
    ]
    return (max(nums) if nums else 0) + 1


def _update_content_types(root, slide_count):
    """更新 [Content_Types].xml 中的 slide Override"""
    # 删除原有的 slide Override
    for override in list(root.findall(f"{{{CT_NS}}}Override")):
        part = override.get("PartName", "")
        if part.startswith("/ppt/slides/slide"):
            root.remove(override)

    # 添加新的 slide Override
    for idx in range(1, slide_count + 1):
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=f"/ppt/slides/slide{idx}.xml",
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )


def _update_presentation_rels(root, slide_count):
    """更新 ppt/_rels/presentation.xml.rels，返回新增 rId 列表"""
    # 统计当前使用的 rId
    existing = [
        rel.get("Id")
        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship")
        if rel.get("Id")
    ]
    start = _next_rid(existing)

    # 删除原有的 slide 关系
    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        target = rel.get("Target", "")
        if target.startswith("slides/slide"):
            root.remove(rel)

    # 添加新的 slide 关系
    new_rel_ids = []
    for idx in range(1, slide_count + 1):
        rid = f"rId{start + idx - 1}"
        new_rel_ids.append(rid)
        ET.SubElement(
            root,
            f"{{{PKG_REL_NS}}}Relationship",
            Id=rid,
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            Target=f"slides/slide{idx}.xml",
        )
    return new_rel_ids


def _update_presentation_xml(root, rel_ids):
    """用新的 rId 顺序重建 p:sldIdLst"""
    ns = {"p": P_NS}
    sld_id_lst = root.find("p:sldIdLst", ns)
    if sld_id_lst is None:
        sld_id_lst = ET.SubElement(root, f"{{{P_NS}}}sldIdLst")
    else:
        for child in list(sld_id_lst):
            sld_id_lst.remove(child)

    # PowerPoint 要求 slideId 从 256 开始
    base = 256
    for idx, rid in enumerate(rel_ids):
        attrib = {f"{{{OD_REL_NS}}}id": rid}
        ET.SubElement(
            sld_id_lst,
            f"{{{P_NS}}}sldId",
            attrib,
            id=str(base + idx),
        )


def duplicate_slides_from_template(template_path, slide_indices, output_path):
    """
    从模板中复制指定的幻灯片，生成新的 PPTX 文件

    :param template_path: 模板 PPTX 文件路径
    :param slide_indices: 要复制的幻灯片索引列表 (从 0 开始)
    :param output_path: 输出 PPTX 文件路径
    """
    # 输入验证
    if not Path(template_path).exists():
        raise FileNotFoundError(f"模板文件不存在: {template_path}")

    if not slide_indices:
        raise ValueError("slide_indices 不能为空")

    # 读取模板 PPTX 的所有文件到内存
    with zipfile.ZipFile(template_path, "r") as tmpl_zip:
        file_bytes = {name: tmpl_zip.read(name) for name in tmpl_zip.namelist()}

    # 提取所有幻灯片和关系文件
    # 注意：PPTX 中的幻灯片编号是 1-based（slide1.xml, slide2.xml...）
    # 但我们的 slide_indices 是 0-based（0, 1, 2...）
    # 所以需要将 PPTX 的编号转换为 0-based 索引
    slide_map = {}
    slide_rel_map = {}
    for name in file_bytes:
        match = SLIDE_RE.fullmatch(name)
        if match:
            pptx_num = int(match.group(1))  # PPTX 中的编号（1-based）
            slide_map[pptx_num - 1] = file_bytes[name]  # 转换为 0-based 索引
        elif name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels"):
            pptx_num = int(re.search(r"slide(\d+)\.xml\.rels", name).group(1))
            slide_rel_map[pptx_num - 1] = file_bytes[name]

    # 验证所有索引都存在
    for idx in slide_indices:
        if idx not in slide_map:
            raise ValueError(
                f"模板中不存在索引为 {idx} 的幻灯片（共 {len(slide_map)} 页）"
            )

    slide_count = len(slide_indices)

    # 更新 presentation.xml.rels
    pres_rels = ET.fromstring(file_bytes["ppt/_rels/presentation.xml.rels"])
    new_rel_ids = _update_presentation_rels(pres_rels, slide_count)

    # 更新 presentation.xml
    pres_xml = ET.fromstring(file_bytes["ppt/presentation.xml"])
    _update_presentation_xml(pres_xml, new_rel_ids)

    # 更新 [Content_Types].xml
    content_types = ET.fromstring(file_bytes["[Content_Types].xml"])
    _update_content_types(content_types, slide_count)

    # 写入新的 PPTX 文件
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
        # 1. 写入所有与 slide 无关的原始文件
        for name, data in file_bytes.items():
            if name.startswith("ppt/slides/slide"):
                continue
            if name.startswith("ppt/slides/_rels/slide"):
                continue
            if name == "[Content_Types].xml":
                out_zip.writestr(
                    name,
                    ET.tostring(content_types, encoding="utf-8", xml_declaration=True),
                )
            elif name == "ppt/_rels/presentation.xml.rels":
                out_zip.writestr(
                    name, ET.tostring(pres_rels, encoding="utf-8", xml_declaration=True)
                )
            elif name == "ppt/presentation.xml":
                out_zip.writestr(
                    name, ET.tostring(pres_xml, encoding="utf-8", xml_declaration=True)
                )
            else:
                out_zip.writestr(name, data)

        # 2. 按顺序写入复制的幻灯片
        for new_idx, source_idx in enumerate(slide_indices, start=1):
            slide_name = f"ppt/slides/slide{new_idx}.xml"
            rel_name = f"ppt/slides/_rels/slide{new_idx}.xml.rels"

            # 写入幻灯片 XML
            out_zip.writestr(slide_name, slide_map[source_idx])

            # 写入关系文件（如果存在）
            if source_idx in slide_rel_map:
                out_zip.writestr(rel_name, slide_rel_map[source_idx])
