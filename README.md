# Script2Slide (S2S)

**自动将 Word 讲稿转换为 PowerPoint 演示文稿**

Script2Slide 是一个基于 Python 的自动化工具，能够读取 Word 讲稿 (`.docx`) 和 PowerPoint 模板 (`.pptx`)，通过 LLM (大语言模型) 智能处理内容，自动生成排版精美的演示文稿。

## 核心特性

- ✅ **内容适应形状**: 通过 LLM 精简内容以适应模板，而非修改模板样式
- ✅ **可插拔 LLM**: 支持多种 LLM Provider (默认 DeepSeek，可选 Gemini)
- ✅ **智能解析**: 自动提取 Word 中的文本、图片、列表和表格
- ✅ **Markdown 支持**: 保留列表格式，自动转换表格为列表
- ✅ **模块化设计**: 遵循 SOLID 原则，代码清晰易维护

## 技术栈

- **Python 3.11+**
- **python-docx**: Word 文档解析
- **python-pptx**: PowerPoint 生成
- **DeepSeek API**: 内容智能处理 (默认，高性价比)
- **Google Gemini API**: 可选的 LLM Provider

## 核心流程

```
Word 讲稿 (.docx)
    ↓
[S1: parse_word.py] ──→ raw_data.json
    ↓
[S2: process_with_llm.py] ──→ data.json
    ↓
[S3: generate_ppt.py] ──→ PowerPoint (.pptx)
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 设置环境变量

```bash
# 使用 DeepSeek（默认，推荐）
export DEEPSEEK_API_KEY="your_deepseek_api_key_here"

# 或者使用 Gemini
export LLM_PROVIDER="gemini"
export GOOGLE_API_KEY="your_google_api_key_here"
```

> 💡 **提示**: DeepSeek 是默认 LLM Provider，价格更实惠，中文效果好。详见 [DeepSeek 设置指南](docs/DEEPSEEK_SETUP.md)

### 3. 准备模板 (一次性设置)

#### 3.1 放置模板文件

将你的 PowerPoint 模板放入 `inputs/templates/` 目录：

```bash
inputs/templates/template.pptx
```

#### 3.2 生成映射文件草稿

使用开发者工具分析模板并生成 `map.json` 草稿：

```bash
python developer_tools/analyze_template.py \
  -t inputs/templates/template.pptx \
  -o config_maps/template.map.json
```

#### 3.3 手动编辑映射文件 ⚠️

**重要**: 打开 `config_maps/template.map.json`，完成以下操作：

1. 审查所有占位符的键名 (如 `title`, `body_text`)，确保语义清晰
2. 将所有 `"AUTO_GENERATED_PLEASE_FILL"` 替换为具体数字
3. 根据模板设置合理的 `max_chars` 和 `max_lines` 约束

示例：

```json
{
  "template_file": "template.pptx",
  "layouts": {
    "18": {
      "name": "标题、正文和图片",
      "layout_index": 18,
      "placeholders": {
        "title": {
          "id": 0,
          "constraints": {
            "max_chars": 30,
            "max_lines": 1
          }
        },
        "body_text": {
          "id": 1,
          "constraints": {
            "max_chars": 200,
            "max_lines": 5
          }
        },
        "picture": {
          "id": 10,
          "constraints": {}
        }
      }
    }
  }
}
```

### 4. 准备讲稿

在 Word 讲稿中使用 `【PPTXX】` 标记来指定布局：

```
【PPT18】
这是第一页的标题

这里是正文内容，可以包含：
- 列表项 1
- 列表项 2
  - 子列表项

【PPT26】
这是第二页的标题
...
```

将讲稿放入 `inputs/scripts/` 目录。

### 5. 执行转换

```bash
python main.py \
  --doc input/scripts/讲稿.docx \
  --template input/templates/template.pptx \
  --map config_maps/template.map.json \
  --output outputs/最终版.pptx
```

### 6. 查看结果

生成的 PowerPoint 文件位于 `outputs/` 目录。

## 项目结构

```
Script2Slide/
│
├── main.py                      # 主入口 (项目启动点)
│
├── input/                       # 输入文件
│   ├── scripts/                 # Word 讲稿
│   └── templates/               # PowerPoint 模板
│
├── config_maps/                 # 模板映射配置
│   └── *.map.json
│
├── core_scripts/                # 核心处理模块
│   ├── parse_word.py           # S1: Word 解析器
│   ├── process_with_llm.py     # S2: LLM 调度器
│   └── generate_ppt.py         # S3: PPT 生成器
│
├── llm_providers/               # LLM Provider 插件
│   ├── base_provider.py        # 抽象接口
│   └── gemini_provider.py      # Gemini 实现
│
├── developer_tools/             # 开发者工具
│   └── analyze_template.py     # 模板分析工具
│
├── outputs/                     # 输出文件
└── requirements.txt             # 依赖列表
```

## 高级配置

### 环境变量

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| `GOOGLE_API_KEY` | Google Gemini API 密钥 | (必需) |
| `LLM_PROVIDER` | LLM Provider 名称 | `gemini` |
| `GEMINI_MODEL` | Gemini 模型名称 | `gemini-2.0-flash-exp` |

### 命令行参数

```bash
python main.py --help
```

主要参数：

- `--doc`: Word 讲稿路径 (必需)
- `--template`: PowerPoint 模板路径 (必需)
- `--map`: 映射文件路径 (必需)
- `--output`: 输出文件路径 (必需)
- `--temp_dir`: 临时目录 (默认: `temp_data`)
- `--keep_temp`: 保留临时文件


## 扩展 LLM Provider

要添加新的 LLM Provider (如 OpenAI)：

1. 在 `llm_providers/` 创建 `openai_provider.py`
2. 继承 `BaseLLMProvider` 并实现 `process_slide` 方法
3. 在 `process_with_llm.py` 的 `LLM_CONFIG` 中添加配置
4. 设置环境变量 `LLM_PROVIDER=openai`

## 常见问题

### Q: 如何调整内容长度限制？

A: 编辑 `config_maps/*.map.json` 中的 `constraints` 字段，调整 `max_chars` 和 `max_lines`。

### Q: 支持哪些图片格式？

A: 支持 Word 中嵌入的所有图片格式，提取后统一保存为 PNG。

### Q: 如何处理表格？

A: Word 中的表格会自动转换为 Markdown 格式，LLM 会将其转换为易读的多级列表。

### Q: 单页处理失败会影响其他页吗？

A: 不会。项目采用防御性编程，单页失败会跳过并继续处理下一页。

## 许可证

MIT License

