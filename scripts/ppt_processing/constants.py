"""PPT 处理常量和命名空间定义。"""

import re
from xml.etree import ElementTree as ET

# 形状名称后缀（用于别名匹配）
SUFFIXES = ("区", "框", "栏")

# 占位符关键词（用于识别文本占位符）
PLACEHOLDER_KEYWORDS = ("文字内容", "字幕", "标题名称", "内容内容")

# 字幕默认文本（需要清除）
SUBTITLE_TEXTS = ("字幕18pt，白色字体深色描边，悬浮阴影。确保在任何底色上都能明确显示",)

# 忽略的形状关键词（装饰性元素）
IGNORE_KEYWORDS = ("背景", "矩形", "圆角", "椭圆", "形状", "图形", "遮罩", "底色")

# 可扩展的形状关键词（标题等）
EXPANDABLE_KEYWORDS = ("标题", "课题", "栏目")

# 手动名称映射（特殊形状名称到实际形状名称）
MANUAL_NAME_MAP = {
    "目录内容区1": ["文本框 9"],
    "目录内容区2": ["文本框 14"],
    "目录内容区3": ["文本框 17"],
    "目录内容区4": ["文本框 20"],
    # 常见字幕框
    "字幕": ["文本框 10", "文本框 32", "文本框 36", "文本框 58", "文本框 121"],
}

# 单位转换常量
EMU_PER_PT = 12700  # EMU per point
H_PADDING = 20000   # 水平内边距（约 1.5 毫米）

# XML 命名空间
NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

# XML 命名空间常量（用于 build_from_json）
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 注册命名空间
ET.register_namespace("", P_NS)
ET.register_namespace("r", OD_REL_NS)

# 正则表达式
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml")
TAG_RE = re.compile(r"ppt/tags/tag(\d+)\.xml")
