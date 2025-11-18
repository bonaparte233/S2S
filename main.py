"""DOCX → JSON → PPT"""

from __future__ import annotations

import argparse
import json
import secrets
from datetime import datetime
from pathlib import Path
from shutil import copyfile

from docx_to_config import generate_config_data
from generate_slides import render_slides


def build_arg_parser() -> argparse.ArgumentParser:
    """构建命令行解析器，暴露给 CLI/GUI 使用。"""
    parser = argparse.ArgumentParser(description="自动读取 DOCX 并生成 PPT")
    parser.add_argument("--docx", required=True, help="讲稿 DOCX 路径")
    parser.add_argument("--template-json", default="template/template.json", help="模板定义 JSON")
    parser.add_argument("--template-list", default="template/template.txt", help="模板编号列表")
    parser.add_argument("--template-ppt", default="template/template.pptx", help="模板 PPTX")
    parser.add_argument("--run-dir", default=None, help="自定义输出目录（默认 temp/run-...）")
    parser.add_argument("--config-name", default="config.json", help="run 目录中的 JSON 名称")
    parser.add_argument("--slides-name", default="slides.pptx", help="run 目录中的 PPT 名称")
    parser.add_argument("--ppt-output", default=None, help="若需单独复制 PPT，请提供完整路径")
    parser.add_argument("--use-llm", action="store_true", help="是否启用大模型")
    parser.add_argument("--llm-provider", default="deepseek", help="大模型提供商标识")
    parser.add_argument("--llm-model", default="deepseek-chat", help="大模型名称")
    parser.add_argument("--course-name", default=None, help="覆盖课程名称")
    parser.add_argument("--college-name", default=None, help="覆盖学院名称")
    parser.add_argument("--lecturer-name", default=None, help="覆盖讲师名称")
    return parser


def _create_pipeline_dir(base: Path = Path("temp")) -> Path:
    """构造 run- 前缀的统一目录，JSON 与 PPT 共享。"""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = secrets.token_hex(2)
    run_dir = base / f"run-{timestamp}-{suffix}"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def run_pipeline(args: argparse.Namespace) -> None:
    """执行完整流程：生成 JSON -> 渲染 PPT -> 输出 run 目录路径。"""
    run_dir = Path(args.run_dir) if args.run_dir else _create_pipeline_dir()
    run_dir.mkdir(parents=True, exist_ok=True)

    overrides = {
        "course": args.course_name,
        "college": args.college_name,
        "lecturer": args.lecturer_name,
    }

    config = generate_config_data(
        docx_path=args.docx,
        template_json=args.template_json,
        template_list=args.template_list,
        use_llm=args.use_llm,
        llm_provider=args.llm_provider,
        llm_model=args.llm_model,
        metadata_overrides=overrides,
        run_dir=run_dir,
    )

    config_path = run_dir / args.config_name
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    slide_result = render_slides(
        template_path=Path(args.template_ppt),
        config=config,
        output_name=args.slides_name,
        run_dir=run_dir,
    )

    if args.ppt_output:
        target = Path(args.ppt_output)
        target.parent.mkdir(parents=True, exist_ok=True)
        copyfile(slide_result["output_path"], target)
        print(f"📄 PPT 已复制到：{target}")

    print(f"✅ JSON：{config_path}")
    print(f"✅ PPT：{slide_result['output_path']}")
    print(f"📁 运行目录：{run_dir}")


def main() -> None:
    args = build_arg_parser().parse_args()
    run_pipeline(args)


if __name__ == "__main__":
    main()
