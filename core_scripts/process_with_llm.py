"""
LLM 调度器 (S2)

功能: 加载 S1 数据，动态加载 LLM 插件，执行内容处理
职责: 只负责调度 LLM，不涉及 Word 解析或 PPT 生成 (SRP 原则)
设计: 依赖抽象 (BaseLLMProvider)，不依赖具体实现 (DIP 原则)
输入: raw_data.json, map.json
输出: data.json
"""

import json
import time
import os
import importlib
from llm_providers.base_provider import BaseLLMProvider


def get_llm_provider(provider_name: str, config: dict) -> BaseLLMProvider:
    """
    动态加载 LLM Provider (工厂模式)

    :param provider_name: Provider 名称 (如 "gemini", "deepseek", "openai")
    :param config: Provider 配置字典
    :return: BaseLLMProvider 实例
    :raises ImportError: 当 Provider 不存在时
    :raises AttributeError: 当 Provider 类不存在时
    """
    # Provider 名称到类名的映射（处理特殊命名）
    CLASS_NAME_MAP = {
        "deepseek": "DeepSeekProvider",
        "gemini": "GeminiProvider",
        "openai": "OpenAIProvider",
    }

    try:
        # 动态导入模块
        module_name = f"llm_providers.{provider_name}_provider"
        module = importlib.import_module(module_name)

        # 获取 Provider 类（使用映射表或默认规则）
        class_name = CLASS_NAME_MAP.get(
            provider_name, f"{provider_name.capitalize()}Provider"
        )
        ProviderClass = getattr(module, class_name)

        # 创建实例
        return ProviderClass(config)

    except ImportError as e:
        raise ImportError(f"无法导入 LLM Provider: {provider_name}。错误: {e}")
    except AttributeError as e:
        raise AttributeError(f"Provider 类 {class_name} 不存在。错误: {e}")


def process_slides_with_llm(
    raw_data_path: str, map_data_path: str, output_path: str = "temp_data/data.json"
) -> str:
    """
    使用 LLM 处理幻灯片内容

    :param raw_data_path: S1 输出的 raw_data.json 路径
    :param map_data_path: 模板映射文件 map.json 路径
    :param output_path: 输出的 data.json 路径
    :return: 输出文件路径
    """
    # 输入验证 (防御性编程)
    if not os.path.exists(raw_data_path):
        raise FileNotFoundError(f"原始数据文件不存在: {raw_data_path}")
    if not os.path.exists(map_data_path):
        raise FileNotFoundError(f"模板映射文件不存在: {map_data_path}")

    # 配置 LLM Provider (配置驱动)
    LLM_PROVIDER_NAME = os.environ.get("LLM_PROVIDER", "deepseek")  # 默认使用 DeepSeek
    LLM_CONFIG = {
        "deepseek": {
            "api_key": os.environ.get("DEEPSEEK_API_KEY"),
            "model_name": os.environ.get("DEEPSEEK_MODEL", "deepseek-chat"),
        },
        "gemini": {
            "api_key": os.environ.get("GOOGLE_API_KEY"),
            "model_name": os.environ.get("GEMINI_MODEL", "gemini-2.0-flash-exp"),
        },
        # 未来可扩展其他 Provider
        # "openai": {...}
    }

    # 加载数据
    with open(raw_data_path, "r", encoding="utf-8") as f:
        raw_slides = json.load(f)

    with open(map_data_path, "r", encoding="utf-8") as f:
        template_map = json.load(f)

    # 初始化 LLM Provider (依赖注入)
    print(f"正在初始化 LLM Provider: {LLM_PROVIDER_NAME}")
    llm_provider = get_llm_provider(
        LLM_PROVIDER_NAME, LLM_CONFIG.get(LLM_PROVIDER_NAME)
    )

    # 遍历处理每一页幻灯片
    final_data = []

    for idx, slide in enumerate(raw_slides, 1):
        layout_id = slide["layout_id"]

        # 检查布局是否在 map.json 中定义
        if layout_id not in template_map["layouts"]:
            print(
                f"⚠️  警告: 讲稿中的 【PPT{layout_id}】 在 map.json 中未定义，已跳过。"
            )
            continue

        layout_info = template_map["layouts"][layout_id]
        placeholders = layout_info["placeholders"]

        # 防御性编程: 对每一页进行错误保护
        try:
            print(f"正在处理 【PPT{layout_id}】 ({idx}/{len(raw_slides)})...")

            # 调用 LLM 处理
            content = llm_provider.process_slide(slide["raw_text"], placeholders)

            # 保存结果
            final_data.append({"layout_id": layout_id, "content": content})

            # 防止 API 速率限制
            time.sleep(1)

        except Exception as e:
            # 优雅失败: 单页失败不影响整体
            print(f"✗ 错误: 处理 【PPT{layout_id}】 失败: {e}。已跳过此页。")
            continue

    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    # 输出 JSON 文件
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(final_data, f, indent=2, ensure_ascii=False)

    print(f"✓ LLM 处理完成，成功处理 {len(final_data)} 页幻灯片")
    print(f"✓ 输出文件: {output_path}")

    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="使用 LLM 处理幻灯片内容")
    parser.add_argument(
        "--raw_data", required=True, help="S1 输出的 raw_data.json 路径"
    )
    parser.add_argument("--map", required=True, help="模板映射文件 map.json 路径")
    parser.add_argument(
        "--output", default="temp_data/data.json", help="输出 data.json 路径"
    )

    args = parser.parse_args()

    try:
        process_slides_with_llm(args.raw_data, args.map, args.output)
    except Exception as e:
        print(f"✗ 错误: {e}")
        exit(1)
