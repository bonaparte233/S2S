"""统一 CLI 入口：导出模板定义 / 生成 JSON / 渲染 PPT / 一键管线。"""

from __future__ import annotations

import argparse
import json
import secrets
from datetime import datetime
from pathlib import Path
from shutil import copyfile

from scripts.docx_to_config import generate_config_data
from scripts.export_template_structure import export_template_structure
from scripts.generate_slides import render_slides


def build_arg_parser() -> argparse.ArgumentParser:
    """构建命令行解析器，供 CLI/GUI 共用。"""
    parser = argparse.ArgumentParser(description="自动化处理模板、讲稿与 PPT")
    parser.add_argument(
        "--mode",
        choices=("template", "script", "slides", "pipeline"),
        default="pipeline",
        help="运行模式：导出模板/仅生成 JSON/仅渲染 PPT/完整流程",
    )
    parser.add_argument("--docx", help="讲稿 DOCX 路径（docx/pipeline 模式必填）")
    parser.add_argument(
        "--template-json", default="template/template.json", help="模板定义 JSON"
    )
    parser.add_argument(
        "--template-list", default="template/template.txt", help="模板编号列表"
    )
    parser.add_argument(
        "--template-ppt", default="template/template.pptx", help="模板 PPTX"
    )
    parser.add_argument(
        "--run-dir", default=None, help="自定义输出目录（默认 temp/run-...）"
    )
    parser.add_argument(
        "--config-name", default="config.json", help="run 目录中的 JSON 名称"
    )
    parser.add_argument(
        "--slides-name", default="slides.pptx", help="run 目录中的 PPT 名称"
    )
    parser.add_argument("--ppt-output", default=None, help="额外复制 PPT 的路径")
    parser.add_argument("--config-output", default=None, help="额外复制 JSON 的路径")
    parser.add_argument("--config-input", help="slides 模式：已有 JSON 文件路径")
    parser.add_argument("--use-llm", action="store_true", help="是否启用大模型")
    parser.add_argument("--llm-provider", default="deepseek", help="大模型提供商标识")
    parser.add_argument("--llm-model", default="deepseek-chat", help="大模型名称")
    parser.add_argument(
        "--llm-base-url", default="http://172.18.75.58:9000", help="大模型接口地址"
    )
    parser.add_argument("--course-name", default=None, help="覆盖课程名称")
    parser.add_argument("--college-name", default=None, help="覆盖学院名称")
    parser.add_argument("--lecturer-name", default=None, help="覆盖讲师名称")
    parser.add_argument(
        "--export-output",
        default="template/exported_template.json",
        help="export-template 模式下导出的 JSON 路径",
    )
    parser.add_argument(
        "--export-mode",
        choices=("semantic", "text"),
        default="semantic",
        help="export-template 模式导出粒度",
    )
    parser.add_argument(
        "--export-include",
        help="export-template 模式：逗号分隔的页码列表（如 1,2,4）",
    )
    return parser


def _create_pipeline_dir(base: Path = Path("temp")) -> Path:
    """构造 run- 前缀的统一目录，JSON 与 PPT 共享。"""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = secrets.token_hex(2)
    run_dir = base / f"run-{timestamp}-{suffix}"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def _require_arg(value, flag: str, mode: str) -> None:
    if value:
        return
    raise SystemExit(f"模式 {mode} 需要提供 {flag}")


def run_export_template(args: argparse.Namespace) -> None:
    """根据 PPT 模板导出 JSON 描述，供 LLM/GUI 使用。"""
    template_path = Path(args.template_ppt)
    include_pages = None
    if args.export_include:
        include_pages = [
            int(item.strip())
            for item in args.export_include.split(",")
            if item.strip().isdigit()
        ]
    data = export_template_structure(template_path, args.export_mode, include_pages)
    output_path = Path(args.export_output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"✅ 模板已导出到：{output_path}")


def run_docx_to_config(args: argparse.Namespace, run_dir: Path | None = None):
    """DOCX→JSON，可单独用于 GUI，也可被管线复用。"""
    mode = args.mode
    _require_arg(args.docx, "--docx", mode)
    base_dir = run_dir or (
        Path(args.run_dir) if args.run_dir else _create_pipeline_dir()
    )
    base_dir.mkdir(parents=True, exist_ok=True)
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
        llm_base_url=args.llm_base_url,
        metadata_overrides=overrides,
        run_dir=base_dir,
    )
    config_path = base_dir / args.config_name
    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if args.config_output:
        target = Path(args.config_output)
        target.parent.mkdir(parents=True, exist_ok=True)
        copyfile(config_path, target)
        print(f"📄 JSON 已复制到：{target}")
    print(f"✅ JSON：{config_path}")
    print(f"📁 运行目录：{base_dir}")
    return config_path, config


def run_generate_slides(
    args: argparse.Namespace,
    run_dir: Path | None = None,
    config_data: dict | None = None,
    config_path: Path | None = None,
):
    """根据 JSON 渲染 PPT，可独立运行。"""
    base_dir = run_dir or (
        Path(args.run_dir) if args.run_dir else _create_pipeline_dir()
    )
    base_dir.mkdir(parents=True, exist_ok=True)

    if config_data is None:
        resolved = config_path or args.config_input
        _require_arg(resolved, "--config-input", args.mode)
        config_path = Path(resolved)
        config_path = config_path.expanduser().resolve()
        config_data = json.loads(config_path.read_text(encoding="utf-8"))
        target = base_dir / args.config_name
        if config_path != target.resolve():
            target.write_text(
                json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            config_path = target
    else:
        config_path = base_dir / args.config_name
        config_path.write_text(
            json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    slide_result = render_slides(
        template_path=Path(args.template_ppt),
        config=config_data,
        output_name=args.slides_name,
        run_dir=base_dir,
    )

    if args.ppt_output:
        target = Path(args.ppt_output)
        target.parent.mkdir(parents=True, exist_ok=True)
        copyfile(slide_result["output_path"], target)
        print(f"📄 PPT 已复制到：{target}")

    print(f"✅ PPT：{slide_result['output_path']}")
    print(f"📁 运行目录：{base_dir}")
    return slide_result["output_path"]


def run_pipeline(args: argparse.Namespace) -> None:
    """完整流程：DOCX → JSON → PPT。"""
    run_dir = Path(args.run_dir) if args.run_dir else _create_pipeline_dir()
    run_dir.mkdir(parents=True, exist_ok=True)
    config_path, config = run_docx_to_config(args, run_dir=run_dir)
    run_generate_slides(
        args, run_dir=run_dir, config_data=config, config_path=config_path
    )


def main() -> None:
    args = build_arg_parser().parse_args()
    if args.mode == "template":
        run_export_template(args)
    elif args.mode == "script":
        run_docx_to_config(args)
    elif args.mode == "slides":
        run_generate_slides(args)
    else:
        run_pipeline(args)


if __name__ == "__main__":
    main()
