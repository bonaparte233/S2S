"""
测试脚本：验证项目设置是否正确

运行此脚本以检查：
1. 所有必需的目录是否存在
2. 所有核心模块是否可以导入
3. 依赖包是否已安装
"""

import os
import sys


def test_directories():
    """测试目录结构"""
    print("检查目录结构...")
    required_dirs = [
        "core_scripts",
        "llm_providers",
        "developer_tools",
        "config_maps",
        "input/scripts",
        "input/templates",
        "outputs",
    ]

    all_exist = True
    for dir_path in required_dirs:
        if os.path.exists(dir_path):
            print(f"  ✓ {dir_path}")
        else:
            print(f"  ✗ {dir_path} (缺失)")
            all_exist = False

    return all_exist


def test_files():
    """测试核心文件"""
    print("\n检查核心文件...")
    required_files = [
        "requirements.txt",
        "main.py",
        "core_scripts/__init__.py",
        "core_scripts/parse_word.py",
        "core_scripts/process_with_llm.py",
        "core_scripts/generate_ppt.py",
        "llm_providers/__init__.py",
        "llm_providers/base_provider.py",
        "llm_providers/gemini_provider.py",
        "developer_tools/analyze_template.py",
    ]

    all_exist = True
    for file_path in required_files:
        if os.path.exists(file_path):
            print(f"  ✓ {file_path}")
        else:
            print(f"  ✗ {file_path} (缺失)")
            all_exist = False

    return all_exist


def test_imports():
    """测试模块导入"""
    print("\n检查模块导入...")

    # 添加 core_scripts 到路径
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "core_scripts"))

    modules_to_test = [
        ("llm_providers.base_provider", "BaseLLMProvider"),
        ("llm_providers.gemini_provider", "GeminiProvider"),
    ]

    all_imported = True
    for module_name, class_name in modules_to_test:
        try:
            module = __import__(module_name, fromlist=[class_name])
            getattr(module, class_name)
            print(f"  ✓ {module_name}.{class_name}")
        except Exception as e:
            print(f"  ✗ {module_name}.{class_name} - {e}")
            all_imported = False

    return all_imported


def test_dependencies():
    """测试依赖包"""
    print("\n检查依赖包...")

    dependencies = [
        ("pptx", "python-pptx"),
        ("docx", "python-docx"),
        ("google.generativeai", "google-generativeai"),
    ]

    all_installed = True
    for import_name, package_name in dependencies:
        try:
            __import__(import_name)
            print(f"  ✓ {package_name}")
        except ImportError:
            print(f"  ✗ {package_name} (未安装)")
            all_installed = False

    return all_installed


def main():
    """主测试函数"""
    print("=" * 60)
    print("Script2Slide 项目设置检查")
    print("=" * 60)

    results = {
        "目录结构": test_directories(),
        "核心文件": test_files(),
        "模块导入": test_imports(),
        "依赖包": test_dependencies(),
    }

    print("\n" + "=" * 60)
    print("检查结果汇总")
    print("=" * 60)

    all_passed = True
    for test_name, passed in results.items():
        status = "✓ 通过" if passed else "✗ 失败"
        print(f"{test_name}: {status}")
        if not passed:
            all_passed = False

    print("=" * 60)

    if all_passed:
        print("\n✓ 所有检查通过！项目设置正确。")
        print("\n下一步:")
        print("1. 设置环境变量: export GOOGLE_API_KEY='your_api_key'")
        print(
            "2. 生成模板映射: python developer_tools/analyze_template.py -t ... -o ..."
        )
        print(
            "3. 运行主程序: python main.py --doc ... --template ... --map ... --output ..."
        )
        return 0
    else:
        print("\n✗ 部分检查失败，请修复上述问题。")
        return 1


if __name__ == "__main__":
    exit(main())
