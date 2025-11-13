"""
Script2Slide 主入口

功能: 串联 S1 (Word 解析), S2 (LLM 处理), S3 (PPT 生成)
职责: 作为总控制器，协调各个模块的执行 (低耦合原则)
"""

import argparse
import os
import shutil

# 从 core_scripts 导入核心模块
from core_scripts.parse_word import parse_word
from core_scripts.process_with_llm import process_slides_with_llm
from core_scripts.generate_ppt import generate_presentation


def main():
    """
    主执行函数
    """
    # 命令行参数解析
    parser = argparse.ArgumentParser(
        description="Script2Slide: 自动将 Word 讲稿转换为 PowerPoint 演示文稿",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python main.py \\
    --doc input/scripts/讲稿.docx \\
    --template input/templates/template.pptx \\
    --map config_maps/template.map.json \\
    --output outputs/最终版.pptx

环境变量:
  DEEPSEEK_API_KEY  DeepSeek API 密钥 (使用 DeepSeek 时必需)
  GOOGLE_API_KEY    Google Gemini API 密钥 (使用 Gemini 时必需)
  LLM_PROVIDER      LLM Provider 名称 (默认: deepseek, 可选: gemini)
  DEEPSEEK_MODEL    DeepSeek 模型名称 (默认: deepseek-chat)
  GEMINI_MODEL      Gemini 模型名称 (默认: gemini-2.0-flash-exp)
        """,
    )

    parser.add_argument(
        "--doc",
        required=True,
        help="Word 讲稿文件路径 (例如: input/scripts/讲稿.docx)",
    )
    parser.add_argument(
        "--template",
        required=True,
        help="PowerPoint 模板文件路径 (例如: input/templates/template.pptx)",
    )
    parser.add_argument(
        "--map",
        required=True,
        help="模板映射文件路径 (例如: config_maps/template.map.json)",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="输出的 PowerPoint 文件路径 (例如: outputs/最终版.pptx)",
    )
    parser.add_argument(
        "--temp_dir", default="temp_data", help="临时数据目录 (默认: temp_data)"
    )
    parser.add_argument(
        "--keep_temp", action="store_true", help="保留临时文件 (默认会清理)"
    )

    args = parser.parse_args()

    # 创建临时目录
    os.makedirs(args.temp_dir, exist_ok=True)

    # 定义临时文件路径
    raw_data_path = os.path.join(args.temp_dir, "raw_data.json")
    data_path = os.path.join(args.temp_dir, "data.json")

    try:
        # ========== S1: 解析 Word 文档 ==========
        print("=" * 60)
        print("步骤 1/3: 解析 Word 文档...")
        print("=" * 60)
        parse_word(args.doc, raw_data_path)
        print(f"✓ Word 解析完成，输出: {raw_data_path}\n")

        # ========== S2: LLM 处理内容 ==========
        print("=" * 60)
        print("步骤 2/3: 使用 LLM 处理内容...")
        print("=" * 60)
        process_slides_with_llm(raw_data_path, args.map, data_path)
        print(f"✓ LLM 处理完成，输出: {data_path}\n")

        # ========== S3: 生成 PowerPoint ==========
        print("=" * 60)
        print("步骤 3/3: 生成 PowerPoint...")
        print("=" * 60)
        generate_presentation(args.template, args.map, data_path, args.output)
        print(f"✓ PowerPoint 生成完成，输出: {args.output}\n")

        # 成功提示
        print("=" * 60)
        print("✓ 所有步骤完成！")
        print("=" * 60)
        print(f"最终文件: {args.output}")

    except Exception as e:
        print(f"\n✗ 错误: {e}")
        raise

    finally:
        # 清理临时文件 (除非用户指定保留)
        if not args.keep_temp and os.path.exists(args.temp_dir):
            print(f"\n清理临时目录: {args.temp_dir}")
            shutil.rmtree(args.temp_dir)
        elif args.keep_temp:
            print(f"\n临时文件保留在: {args.temp_dir}")


if __name__ == "__main__":
    main()
