"""连接器处理工具。"""

import zipfile
from pathlib import Path
from typing import Dict, List

from lxml import etree

from scripts.ppt_processing.constants import SLIDE_RE, NSMAP


def extract_connectors(pptx_path: Path) -> Dict[int, List[bytes]]:
    """从 PPT 中提取连接器元素。

    Args:
        pptx_path: PPT 文件路径

    Returns:
        幻灯片索引到连接器 XML 字节列表的映射
    """
    connectors = {}
    with zipfile.ZipFile(pptx_path, "r") as zf:
        for name in zf.namelist():
            match = SLIDE_RE.fullmatch(name)
            if not match:
                continue
            slide_idx = int(match.group(1))
            root = etree.fromstring(zf.read(name))
            nodes = root.xpath(".//p:cxnSp", namespaces=NSMAP)
            if nodes:
                connectors[slide_idx] = [etree.tostring(node) for node in nodes]
    return connectors


def restore_connectors(
    pptx_path: Path,
    connectors: Dict[int, List[bytes]]
) -> None:
    """将连接器元素恢复到 PPT 中。

    Args:
        pptx_path: PPT 文件路径
        connectors: 幻灯片索引到连接器 XML 字节列表的映射
    """
    if not connectors:
        return

    with zipfile.ZipFile(pptx_path, "r") as src:
        entries = {name: src.read(name) for name in src.namelist()}

    modified = False
    for name, data in list(entries.items()):
        match = SLIDE_RE.fullmatch(name)
        if not match:
            continue
        slide_idx = int(match.group(1))
        snippets = connectors.get(slide_idx)
        if not snippets:
            continue
        root = etree.fromstring(data)
        sp_tree = root.find(".//p:spTree", namespaces=NSMAP)
        if sp_tree is None:
            continue
        # 移除现有连接器
        for node in sp_tree.findall("p:cxnSp", namespaces=NSMAP):
            sp_tree.remove(node)
        # 添加保存的连接器
        for snippet in snippets:
            sp_tree.append(etree.fromstring(snippet))
        entries[name] = etree.tostring(root, encoding="utf-8", xml_declaration=True)
        modified = True

    if not modified:
        return

    with zipfile.ZipFile(pptx_path, "w") as dst:
        for name, data in entries.items():
            dst.writestr(name, data)
