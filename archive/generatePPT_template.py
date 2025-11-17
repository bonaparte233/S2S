"""根据 JSON 描述，直接在 PPTX 包层面复制模板页并重新排序。

核心思路：
1. 不通过 python-pptx API 操作幻灯片内容，而是直接处理 zip 中的 XML/关系文件。
2. 复制 slide、关系（rels）、tag 元数据与图片，保证生成的 PPT 与模板完全一致。
3. 根据 JSON 的顺序，重写 presentation.xml 与 _rels 文件，使 PowerPoint 认为这是原生输出。
"""

import argparse
import json
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# --- 全局常量：描述需要处理的文件模式与命名空间 ---
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml")
TAG_RE = re.compile(r"ppt/tags/tag(\d+)\.xml")
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

ET.register_namespace("", P_NS)
ET.register_namespace("r", OD_REL_NS)


def _next_rid(existing_ids):
    """生成下一个未被占用的 rId 序号"""
    nums = [
        int(rid[3:])
        for rid in existing_ids
        if rid.startswith("rId") and rid[3:].isdigit()
    ]
    return (max(nums) if nums else 0) + 1


def _update_content_types(root, slide_count, new_tag_parts):
    """更新 [Content_Types].xml 中的 slide Override，并追加 tag 定义"""
    # 先删掉原有的 slide Override，避免旧顺序影响新 PPT
    for override in list(root.findall(f"{{{CT_NS}}}Override")):
        part = override.get("PartName", "")
        if part.startswith("/ppt/slides/slide"):
            root.remove(override)

    for idx in range(1, slide_count + 1):
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=f"/ppt/slides/slide{idx}.xml",
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )

    for part_name in new_tag_parts:
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=f"/{part_name}",
            ContentType="application/vnd.ms-powerpoint.tags+xml",
        )


def _update_presentation_rels(root, slide_count):
    """更新 ppt/_rels/presentation.xml.rels，返回新增 rId 列表"""
    # 统计当前文件使用的 rId，生成下一个可用序号，避免与模板原数据冲突
    existing = [
        rel.get("Id")
        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship")
        if rel.get("Id")
    ]
    start = _next_rid(existing)

    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        target = rel.get("Target", "")
        if target.startswith("slides/slide"):
            root.remove(rel)

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

    # PowerPoint 要求 slideId 从一个固定值开始，这里沿用 256 起步的策略
    base = 256
    for idx, rid in enumerate(rel_ids):
        attrib = {f"{{{OD_REL_NS}}}id": rid}
        ET.SubElement(
            sld_id_lst,
            f"{{{P_NS}}}sldId",
            attrib,
            id=str(base + idx),
        )


def build_from_json(template_path, json_path, output_path):
    """复制原始模板 pptx，并按 JSON 顺序重新组织 slide 文件"""
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    pages = data.get("ppt_pages", [])
    if not pages:
        raise ValueError("JSON 中未找到 ppt_pages 内容")

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_copy = Path(tmpdir) / "working.pptx"
        shutil.copyfile(template_path, temp_copy)

        # 将模板 PPT 的所有文件读入内存，便于自由重写
        with zipfile.ZipFile(template_path, "r") as tmpl_zip:
            file_bytes = {name: tmpl_zip.read(name) for name in tmpl_zip.namelist()}

        slide_map = {}
        slide_rel_map = {}
        for name in file_bytes:
            match = SLIDE_RE.fullmatch(name)
            if match:
                slide_map[int(match.group(1))] = file_bytes[name]
            elif name.startswith("ppt/slides/_rels/slide") and name.endswith(
                ".xml.rels"
            ):
                num = int(re.search(r"slide(\d+)\.xml\.rels", name).group(1))
                slide_rel_map[num] = file_bytes[name]

        slide_count = len(slide_map)
        if slide_count == 0:
            raise ValueError("模板中未找到任何 slide 文件")

        tag_nums = [
            int(m.group(1))
            for name in file_bytes
            if (m := TAG_RE.fullmatch(name)) is not None
        ]
        next_tag_num = max(tag_nums) if tag_nums else 0
        extra_tag_parts = []
        extra_tag_files = {}

        # 根据 JSON 记录应使用的模板页编号，同时打印处理进度
        selected_slides = []
        for idx, page in enumerate(pages, start=1):
            tmpl_num = page.get("template_page_num")
            page_type = page.get("page_type", "未知版式")
            if tmpl_num is None:
                raise ValueError(f"第{idx}条缺少 template_page_num")
            if tmpl_num not in slide_map:
                raise ValueError(f"模板中不存在第{tmpl_num}页（来自第{idx}条 {page_type}）")
            selected_slides.append((tmpl_num, page_type))
            print(f"✅ 生成第{idx}页：{page_type}（模板第{tmpl_num}页）")

        slide_total = len(selected_slides)

        pres_rels = ET.fromstring(file_bytes["ppt/_rels/presentation.xml.rels"])
        new_rel_ids = _update_presentation_rels(pres_rels, slide_total)

        pres_xml = ET.fromstring(file_bytes["ppt/presentation.xml"])
        _update_presentation_xml(pres_xml, new_rel_ids)

        def clone_tags(rel_bytes):
            nonlocal next_tag_num
            if not rel_bytes:
                return rel_bytes
            rel_tree = ET.fromstring(rel_bytes)
            for rel in rel_tree.findall(f"{{{PKG_REL_NS}}}Relationship"):
                if (
                    rel.get("Type")
                    != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
                ):
                    continue
                target = rel.get("Target")
                canonical = posixpath.normpath(posixpath.join("ppt/slides", target))
                if canonical not in file_bytes:
                    continue
                # tag 关系在 PPT 中要求唯一，因此为每条关系生成新的 tag 文件
                next_tag_num += 1
                new_part = f"ppt/tags/tag{next_tag_num}.xml"
                rel.set("Target", posixpath.relpath(new_part, "ppt/slides"))
                extra_tag_parts.append(new_part)
                extra_tag_files[new_part] = file_bytes[canonical]
            return ET.tostring(rel_tree, encoding="utf-8", xml_declaration=True)

        prepared_rel_bytes = {}
        for idx, (tmpl_num, _) in enumerate(selected_slides, start=1):
            if tmpl_num in slide_rel_map:
                prepared_rel_bytes[idx] = clone_tags(slide_rel_map[tmpl_num])

        content_types = ET.fromstring(file_bytes["[Content_Types].xml"])
        _update_content_types(content_types, slide_total, extra_tag_parts)

        with zipfile.ZipFile(output_path, "w") as out_zip:
            # 1. 先写入所有与 slide 无关的原始文件（主题、字体、媒体等）
            for name, data in file_bytes.items():
                if name.startswith("ppt/slides/slide"):
                    continue
                if name.startswith("ppt/slides/_rels/slide"):
                    continue
                if name == "[Content_Types].xml":
                    out_zip.writestr(
                        name, ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
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

            # 2. 把 JSON 指定顺序的 slide 与关系文件依次写入
            for idx, (tmpl_num, _) in enumerate(selected_slides, start=1):
                slide_name = f"ppt/slides/slide{idx}.xml"
                rel_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
                out_zip.writestr(slide_name, slide_map[tmpl_num])
                rel_bytes = prepared_rel_bytes.get(idx)
                if rel_bytes:
                    out_zip.writestr(rel_name, rel_bytes)

            # 3. 写入为 tag 关系复制的新文件，确保 PPT 打开不会再修复
            for name, data in extra_tag_files.items():
                out_zip.writestr(name, data)

    print(f"\n🎉 新 PPT 输出完成：{output_path}")


def main():
    parser = argparse.ArgumentParser(description="根据 JSON 顺序复制模板页")
    parser.add_argument("--template", required=True, help="模板 PPTX 路径")
    parser.add_argument("--json", required=True, help="输入 JSON 文件")
    parser.add_argument("--output", default="generated_template.pptx", help="输出 PPTX")
    args = parser.parse_args()

    build_from_json(args.template, args.json, args.output)


if __name__ == "__main__":
    main()
