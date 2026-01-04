"""PPT 处理子包 - 将 generate_slides.py 拆分为多个模块。

模块结构：
- constants.py: 常量和命名空间定义
- xml_utils.py: XML/Rels 文件处理
- shape_utils.py: 形状检测和遍历工具
- text_utils.py: 文本格式化和填充
- image_utils.py: 图片替换逻辑
- layout_utils.py: 布局调整和自适应
- slide_builder.py: 幻灯片构建和填充
- connector_utils.py: 连接器处理
"""

from scripts.ppt_processing.constants import (
    NSMAP,
    P_NS,
    PKG_REL_NS,
    OD_REL_NS,
    CT_NS,
    SLIDE_RE,
    TAG_RE,
    SUFFIXES,
    PLACEHOLDER_KEYWORDS,
    SUBTITLE_TEXTS,
    IGNORE_KEYWORDS,
    EXPANDABLE_KEYWORDS,
    MANUAL_NAME_MAP,
    EMU_PER_PT,
    H_PADDING,
)
from scripts.ppt_processing.xml_utils import (
    clean_rels_namespace,
    next_rid,
    update_content_types,
    update_presentation_rels,
    update_presentation_xml,
)
from scripts.ppt_processing.shape_utils import (
    iter_shapes,
    iter_shapes_with_path,
    is_picture_shape,
    is_placeholder_shape,
    shape_aliases,
    shape_tags,
    detect_prefix,
    candidate_keys,
    take_shape,
)
from scripts.ppt_processing.text_utils import (
    copy_run_format,
    copy_para_format,
    set_shape_text,
)
from scripts.ppt_processing.image_utils import (
    replace_picture,
    safe_remove_shape,
)
from scripts.ppt_processing.layout_utils import (
    apply_layout_rules,
    clear_default_subtitles,
)
from scripts.ppt_processing.slide_builder import (
    build_from_json,
    fill_slide,
    flatten_content,
    delete_empty_group_shapes,
    render_slides,
    create_run_dir,
)
from scripts.ppt_processing.connector_utils import (
    extract_connectors,
    restore_connectors,
)

__all__ = [
    # Constants
    "NSMAP",
    "P_NS",
    "PKG_REL_NS",
    "OD_REL_NS",
    "CT_NS",
    "SLIDE_RE",
    "TAG_RE",
    "SUFFIXES",
    "PLACEHOLDER_KEYWORDS",
    "SUBTITLE_TEXTS",
    "IGNORE_KEYWORDS",
    "EXPANDABLE_KEYWORDS",
    "MANUAL_NAME_MAP",
    "EMU_PER_PT",
    "H_PADDING",
    # XML Utils
    "clean_rels_namespace",
    "next_rid",
    "update_content_types",
    "update_presentation_rels",
    "update_presentation_xml",
    # Shape Utils
    "iter_shapes",
    "iter_shapes_with_path",
    "is_picture_shape",
    "is_placeholder_shape",
    "shape_aliases",
    "shape_tags",
    "detect_prefix",
    "candidate_keys",
    "take_shape",
    # Text Utils
    "copy_run_format",
    "copy_para_format",
    "set_shape_text",
    # Image Utils
    "replace_picture",
    "safe_remove_shape",
    # Layout Utils
    "apply_layout_rules",
    "clear_default_subtitles",
    # Slide Builder
    "build_from_json",
    "fill_slide",
    "flatten_content",
    "delete_empty_group_shapes",
    "render_slides",
    "create_run_dir",
    # Connector Utils
    "extract_connectors",
    "restore_connectors",
]
