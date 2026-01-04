"""常量和正则表达式定义。"""

from __future__ import annotations

import os
import re

# PPT 标记正则
MARKER_RE = re.compile(r"【PPT(\d+)】")

# 图片文件名模板
IMAGE_NAME_TEMPLATE = "doc_image_{idx}.{ext}"

# 调试标志
DEBUG_LLM = os.getenv("DEBUG_LLM", "false").lower() in ("true", "1", "yes")

# 章节标题正则表达式（只追踪两级）
# 一级标题（章节）：汉字序号，如 一、xxx / 第一章 xxx
HEADING_L1_RE = re.compile(
    r"^(?:"
    r"[一二三四五六七八九十]+、|"
    r"第[一二三四五六七八九十]+[章节部分]"
    r")(.+)",
    re.MULTILINE,
)

# 二级标题（知识点）：数字序号，如 1. xxx / 1.1 xxx / （1）xxx
HEADING_L2_RE = re.compile(
    r"^(?:"
    r"\d+\.\s*|"
    r"\d+\.\d+\s*|"
    r"[（\(]\d+[）\)]"
    r")(.+)",
    re.MULTILINE,
)
