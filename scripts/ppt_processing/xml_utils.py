"""XML 和 Rels 文件处理工具。"""

import re
from typing import List
from xml.etree import ElementTree as ET

from scripts.ppt_processing.constants import PKG_REL_NS, OD_REL_NS, CT_NS, P_NS


def clean_rels_namespace(xml_bytes: bytes) -> bytes:
    """清理 rels XML 中的 ns0: 前缀，PowerPoint 对此敏感。

    ElementTree 在输出时会给未注册的命名空间添加 ns0: 前缀，
    但 PowerPoint 期望 rels 文件使用无前缀的默认命名空间。

    Args:
        xml_bytes: 原始 XML 字节数据

    Returns:
        清理后的 XML 字节数据
    """
    xml_str = xml_bytes.decode("utf-8")
    # 替换 ns0: 前缀和 xmlns:ns0 声明
    xml_str = xml_str.replace("ns0:", "").replace(":ns0", "")
    # 处理可能的 ns1:, ns2: 等
    xml_str = re.sub(r"\bns\d+:", "", xml_str)
    xml_str = re.sub(r":ns\d+\b", "", xml_str)
    return xml_str.encode("utf-8")


def next_rid(existing_ids: List[str]) -> int:
    """生成下一个未被占用的 rId 序号。

    Args:
        existing_ids: 已存在的 rId 列表

    Returns:
        下一个可用的 rId 序号
    """
    nums = [
        int(rid[3:])
        for rid in existing_ids
        if rid.startswith("rId") and rid[3:].isdigit()
    ]
    return (max(nums) if nums else 0) + 1


def update_content_types(
    root: ET.Element,
    slide_count: int,
    new_tag_parts: List[str]
) -> None:
    """更新 [Content_Types].xml 中的 slide Override，并追加 tag 定义。

    Args:
        root: Content_Types XML 根元素
        slide_count: 幻灯片总数
        new_tag_parts: 新增的 tag 文件路径列表
    """
    # 删除原有的 slide Override 和 notes 相关条目
    for override in list(root.findall(f"{{{CT_NS}}}Override")):
        part = override.get("PartName", "")
        if part.startswith("/ppt/slides/slide"):
            root.remove(override)
        if "/notesSlides/" in part or "/notesMasters/" in part:
            root.remove(override)

    # 添加新的 slide Override
    for idx in range(1, slide_count + 1):
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=f"/ppt/slides/slide{idx}.xml",
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )

    # 添加 tag Override
    for part_name in new_tag_parts:
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=f"/{part_name}",
            ContentType="application/vnd.ms-powerpoint.tags+xml",
        )


def update_presentation_rels(root: ET.Element, slide_count: int) -> List[str]:
    """更新 ppt/_rels/presentation.xml.rels，返回新增 rId 列表。

    Args:
        root: presentation.xml.rels XML 根元素
        slide_count: 幻灯片总数

    Returns:
        新增的 rId 列表
    """
    # 统计当前文件使用的 rId
    existing = [
        rel.get("Id")
        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship")
        if rel.get("Id")
    ]
    start = next_rid(existing)

    # 删除原有 slide 和 notes 引用
    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        target = rel.get("Target", "")
        if target.startswith("slides/slide"):
            root.remove(rel)
        if target.startswith("notesSlides/") or target.startswith("notesMasters/"):
            root.remove(rel)

    # 添加新的 slide 引用
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


def update_presentation_xml(root: ET.Element, rel_ids: List[str]) -> None:
    """用新的 rId 顺序重建 p:sldIdLst，并删除 notesMasterIdLst。

    Args:
        root: presentation.xml XML 根元素
        rel_ids: 新的 rId 列表
    """
    ns = {"p": P_NS}
    sld_id_lst = root.find("p:sldIdLst", ns)
    if sld_id_lst is None:
        sld_id_lst = ET.SubElement(root, f"{{{P_NS}}}sldIdLst")
    else:
        for child in list(sld_id_lst):
            sld_id_lst.remove(child)

    # PowerPoint 要求 slideId 从 256 起步
    base = 256
    for idx, rid in enumerate(rel_ids):
        attrib = {f"{{{OD_REL_NS}}}id": rid}
        ET.SubElement(
            sld_id_lst,
            f"{{{P_NS}}}sldId",
            attrib,
            id=str(base + idx),
        )

    # 删除 notesMasterIdLst
    notes_master_lst = root.find("p:notesMasterIdLst", ns)
    if notes_master_lst is not None:
        root.remove(notes_master_lst)
