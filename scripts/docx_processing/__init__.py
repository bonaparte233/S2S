"""DOCX 处理子包 - 将 docx_to_config.py 拆分为多个模块。

模块结构：
- constants.py: 常量和正则表达式
- docx_parser.py: DOCX 文件解析
- docx_table_parser.py: DOCX 表格解析
- template_utils.py: 模板定义处理
- json_utils.py: JSON 解析工具
- llm_prompts.py: LLM Prompt 构建
- llm_processor.py: LLM 调用逻辑
- slide_filler.py: 幻灯片内容填充
- special_pages.py: 特殊页面（封面、目录、结束页）
"""

from scripts.docx_processing.constants import (
    MARKER_RE,
    IMAGE_NAME_TEMPLATE,
    DEBUG_LLM,
    HEADING_L1_RE,
    HEADING_L2_RE,
)
from scripts.docx_processing.docx_parser import (
    parse_docx_blocks,
    SectionContext,
)
from scripts.docx_processing.docx_table_parser import (
    extract_tables_from_docx,
)
from scripts.docx_processing.template_utils import (
    load_template_defs,
    clone_schema,
    assign_in_schema,
    get_in_schema,
    strip_values,
    empty_content,
)
from scripts.docx_processing.json_utils import (
    ensure_json_object,
    ensure_json_array,
)
from scripts.docx_processing.llm_prompts import (
    build_fill_prompt,
    build_multimodal_messages,
    encode_image,
)
from scripts.docx_processing.llm_processor import (
    choose_llm,
    llm_fill_slide,
    llm_plan_slides,
    llm_preprocess_script,
)
from scripts.docx_processing.slide_filler import (
    fill_by_markers,
    preprocess_and_fill,
)
from scripts.docx_processing.special_pages import (
    prepend_cover_page,
    insert_toc_page,
    append_end_page,
)

__all__ = [
    # Constants
    "MARKER_RE",
    "IMAGE_NAME_TEMPLATE",
    "DEBUG_LLM",
    "HEADING_L1_RE",
    "HEADING_L2_RE",
    # DOCX Parser
    "parse_docx_blocks",
    "SectionContext",
    # DOCX Table Parser
    "extract_tables_from_docx",
    # Template Utils
    "load_template_defs",
    "clone_schema",
    "assign_in_schema",
    "get_in_schema",
    "strip_values",
    "empty_content",
    # JSON Utils
    "ensure_json_object",
    "ensure_json_array",
    # LLM Prompts
    "build_fill_prompt",
    "build_multimodal_messages",
    "encode_image",
    # LLM Processor
    "choose_llm",
    "llm_fill_slide",
    "llm_plan_slides",
    "llm_preprocess_script",
    # Slide Filler
    "fill_by_markers",
    "preprocess_and_fill",
    # Special Pages
    "prepend_cover_page",
    "insert_toc_page",
    "append_end_page",
]
