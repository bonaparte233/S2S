"""根据 DOCX 讲稿和模板定义生成 JSON，可选调用 LLM。

该模块是 DOCX 处理的主入口，核心逻辑已拆分到 scripts/docx_processing/ 子包。
"""

from __future__ import annotations

import argparse
import json
import secrets
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

# 从子模块导入核心功能
from scripts.docx_processing import (
    parse_docx_blocks,
    load_template_defs,
    choose_llm,
    preprocess_and_fill,
    fill_by_markers,
    prepend_cover_page,
    insert_toc_page,
    append_end_page,
)
from scripts.docx_processing.template_utils import strip_values
from scripts.docx_processing.slide_filler import _parse_preprocessed_script

# 保持向后兼容的导出
from scripts.docx_processing import (
    MARKER_RE,
    IMAGE_NAME_TEMPLATE,
    DEBUG_LLM,
    SectionContext,
    llm_fill_slide,
    llm_plan_slides,
    llm_preprocess_script,
    ensure_json_object,
    ensure_json_array,
)

__all__ = [
    "generate_config_data",
    "process_docx",
    "parse_docx_blocks",
    "load_template_defs",
    "choose_llm",
    "MARKER_RE",
    "IMAGE_NAME_TEMPLATE",
    "DEBUG_LLM",
    "SectionContext",
    "llm_fill_slide",
    "llm_plan_slides",
    "llm_preprocess_script",
]


def _create_run_dir(base_dir: Path = Path("temp")) -> Path:
    """创建带时间戳前缀的运行目录。"""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = secrets.token_hex(2)
    run_dir = base_dir / f"script-{timestamp}-{suffix}"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def generate_config_data(
    docx_path: str,
    template_json: str,
    template_list: str,
    use_llm: bool,
    llm_provider: str,
    llm_model: Optional[str],
    llm_base_url: Optional[str],
    metadata_overrides: Optional[Dict[str, str]],
    run_dir: Path,
    user_prompt: Optional[str] = None,
) -> Dict:
    """核心逻辑：生成 JSON 内容，供 GUI/CLI 复用。

    Args:
        docx_path: DOCX 文件路径
        template_json: 模板定义 JSON 文件路径
        template_list: 模板编号列表文件路径（已废弃）
        use_llm: 是否启用 LLM
        llm_provider: LLM 提供商
        llm_model: LLM 模型名称
        llm_base_url: LLM 接口地址
        metadata_overrides: 元数据覆盖
        run_dir: 运行目录
        user_prompt: 用户自定义提示

    Returns:
        生成的配置数据字典
    """
    metadata_overrides = metadata_overrides or {}
    image_dir = run_dir / "images"
    blocks, has_marker, metadata = parse_docx_blocks(docx_path, image_dir)

    # 应用元数据覆盖
    for key in ("course", "college", "lecturer"):
        if metadata_overrides.get(key):
            metadata[key] = metadata_overrides[key]

    templates, global_config = load_template_defs(template_json, template_list)
    llm = choose_llm(use_llm, llm_provider, llm_model, llm_base_url)

    # 统一走预处理流程
    pages, chapters = preprocess_and_fill(
        blocks, templates, llm, metadata, user_prompt, run_dir, has_marker, global_config
    )

    if not pages:
        raise ValueError("未生成任何幻灯片内容，请检查讲稿或模板。")

    # 添加特殊页面
    prepend_cover_page(pages, templates, metadata, llm, global_config)
    insert_toc_page(pages, templates, metadata, chapters, global_config)
    append_end_page(pages, templates, metadata, global_config)

    # 清理输出格式
    stripped_pages = []
    for page in pages:
        stripped_pages.append({
            "page_type": page.get("page_type"),
            "template_page_num": page.get("template_page_num"),
            "content": strip_values(page.get("content", {})),
        })

    return {"ppt_pages": stripped_pages}


def process_docx(
    docx_path: str,
    template_json: str,
    template_list: str,
    output_path: Optional[str],
    use_llm: bool,
    llm_provider: str,
    llm_model: Optional[str],
    llm_base_url: Optional[str],
    override_course: Optional[str],
    override_college: Optional[str],
    override_lecturer: Optional[str],
    run_dir: Optional[str],
    config_name: str,
) -> None:
    """CLI 包装：处理参数、保证 run 目录存在，并额外复制文件到 output。"""
    metadata_overrides = {
        "course": override_course,
        "college": override_college,
        "lecturer": override_lecturer,
    }

    base_dir = Path(run_dir) if run_dir else _create_run_dir()
    base_dir.mkdir(parents=True, exist_ok=True)
    config_path = base_dir / config_name

    config = generate_config_data(
        docx_path,
        template_json,
        template_list,
        use_llm,
        llm_provider,
        llm_model,
        llm_base_url,
        metadata_overrides,
        base_dir,
    )

    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if output_path:
        explicit = Path(output_path)
        explicit.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(config_path, explicit)
        print(f"📄 另存为：{explicit}")

    print(f"✅ 已生成 JSON：{config_path}")
    print(f"📁 资源输出目录：{base_dir}")


def continue_from_preprocessed(
    preprocessed_path: str,
    template_json: str,
    template_pptx: str,
    use_llm: bool = True,
    llm_provider: str = "deepseek",
    llm_model: Optional[str] = None,
    llm_base_url: Optional[str] = None,
    override_course: Optional[str] = None,
    override_college: Optional[str] = None,
    override_lecturer: Optional[str] = None,
    run_dir: Optional[str] = None,
    config_name: str = "config.json",
) -> Dict:
    """从预处理后的讲稿继续生成 config.json。"""
    preprocessed_file = Path(preprocessed_path)
    if not preprocessed_file.exists():
        raise FileNotFoundError(f"预处理讲稿不存在: {preprocessed_path}")

    run_dir_path = Path(run_dir) if run_dir else preprocessed_file.parent
    image_dir = run_dir_path / "images"

    preprocessed_text = preprocessed_file.read_text(encoding="utf-8")
    blocks = _parse_preprocessed_script(preprocessed_text, image_dir)
    print(f"✅ 解析到 {len(blocks)} 个 blocks")

    templates, global_config = load_template_defs(template_json, None)
    print(f"✅ 加载了 {len(templates)} 个模板页")

    llm = choose_llm(use_llm, llm_provider, llm_model, llm_base_url)

    metadata = {
        "course": override_course or "",
        "college": override_college or "",
        "lecturer": override_lecturer or "",
    }

    print("🔄 开始填充内容...")
    pages, chapters = fill_by_markers(
        blocks, templates, llm, metadata, user_prompt=None, global_config=global_config
    )
    print(f"✅ 生成了 {len(pages)} 页")

    prepend_cover_page(pages, templates, metadata, llm, global_config)
    insert_toc_page(pages, templates, metadata, chapters, global_config)
    append_end_page(pages, templates, metadata, global_config)

    config = {
        "template": template_pptx,
        "pages": pages,
    }

    config_path = run_dir_path / config_name
    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"✅ 已保存到 {config_path}")

    return config


def build_arg_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""
    parser = argparse.ArgumentParser(description="根据 DOCX 讲稿生成 PPT 配置 JSON。")
    parser.add_argument("--docx", default=None, help="讲稿 DOCX 路径")
    parser.add_argument(
        "--from-preprocessed", default=None,
        help="从预处理后的讲稿(.md)继续，跳过预处理步骤",
    )
    parser.add_argument(
        "--template-json", default="template/template.json",
        help="模板定义 JSON 文件",
    )
    parser.add_argument(
        "--template-list", default="template/template.txt",
        help="可用模板编号列表 txt",
    )
    parser.add_argument(
        "--output", default=None,
        help="如需额外复制一份 JSON，请提供完整路径",
    )
    parser.add_argument("--use-llm", action="store_true", help="启用大模型填充/排版")
    parser.add_argument("--llm-provider", default="deepseek", help="大模型提供商")
    parser.add_argument("--llm-model", default="deepseek-chat", help="大模型名称")
    parser.add_argument(
        "--llm-base-url", default="http://172.18.75.58:9000",
        help="自定义大模型接口地址",
    )
    parser.add_argument("--course-name", default=None, help="手动指定课程/项目名称")
    parser.add_argument("--college-name", default=None, help="手动指定学院/单位")
    parser.add_argument("--lecturer-name", default=None, help="手动指定主讲教师姓名")
    parser.add_argument("--run-dir", default=None, help="指定输出目录")
    parser.add_argument("--config-name", default="config.json", help="配置文件名称")
    parser.add_argument("--template-pptx", default=None, help="模板 PPTX 路径")
    return parser


def main() -> None:
    """CLI 主入口。"""
    args = build_arg_parser().parse_args()

    if args.from_preprocessed:
        if not args.template_pptx:
            print("❌ 使用 --from-preprocessed 时必须指定 --template-pptx")
            return
        continue_from_preprocessed(
            preprocessed_path=args.from_preprocessed,
            template_json=args.template_json,
            template_pptx=args.template_pptx,
            use_llm=args.use_llm,
            llm_provider=args.llm_provider,
            llm_model=args.llm_model,
            llm_base_url=args.llm_base_url,
            override_course=args.course_name,
            override_college=args.college_name,
            override_lecturer=args.lecturer_name,
            run_dir=args.run_dir,
            config_name=args.config_name,
        )
        return

    if not args.docx:
        print("❌ 必须指定 --docx 或 --from-preprocessed")
        return

    process_docx(
        docx_path=args.docx,
        template_json=args.template_json,
        template_list=args.template_list,
        output_path=args.output,
        use_llm=args.use_llm,
        llm_provider=args.llm_provider,
        llm_model=args.llm_model,
        llm_base_url=args.llm_base_url,
        override_course=args.course_name,
        override_college=args.college_name,
        override_lecturer=args.lecturer_name,
        run_dir=args.run_dir,
        config_name=args.config_name,
    )


if __name__ == "__main__":
    main()
