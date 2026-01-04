"""根据 JSON 描述复制模板、填充文本/图片并生成最终 PPT。

此模块已重构为模块化结构，核心逻辑位于 scripts/ppt_processing/ 子包中。
本文件保留为向后兼容的入口点和 CLI 接口。
"""

import argparse
import json
import shutil
from pathlib import Path
from typing import Dict, Optional

# 从子模块导入核心功能
from scripts.ppt_processing import (
    # 主要入口函数
    render_slides,
    build_from_json,
    fill_slide,
    create_run_dir,
    # 工具函数（向后兼容导出）
    flatten_content,
    delete_empty_group_shapes,
    # 常量（向后兼容导出）
    NSMAP,
    P_NS,
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

# 向后兼容：导出带下划线前缀的内部函数名
_create_run_dir = create_run_dir
_fill_slide = fill_slide
_flatten_content = flatten_content
_delete_empty_group_shapes = delete_empty_group_shapes


def main():
    """CLI 入口函数。"""
    parser = argparse.ArgumentParser(description="读取 JSON 并填充模板生成 PPT")
    parser.add_argument("--template", required=True, help="模板 PPTX 路径")
    parser.add_argument("--json", required=True, help="描述内容的 JSON 文件")
    parser.add_argument(
        "--output", default="final_output.pptx", help="输出 PPTX 文件名或路径"
    )
    parser.add_argument(
        "--run-dir", default=None, help="输出 run 目录（默认 temp/run-...）"
    )
    args = parser.parse_args()

    config = json.loads(Path(args.json).read_text(encoding="utf-8"))
    run_dir = Path(args.run_dir) if args.run_dir else None
    result = render_slides(
        Path(args.template), config, Path(args.output).name, run_dir
    )
    final_path = result["output_path"]

    if Path(args.output).is_absolute():
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(final_path, Path(args.output))
        print(f"📄 另存为：{args.output}")

    print(f"🎯 已根据内容生成 PPT：{final_path}")


if __name__ == "__main__":
    main()
