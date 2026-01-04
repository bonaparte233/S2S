# S2S 后端 API 文档

本目录包含 S2S (Script to Slides) 的核心后端模块，供第三方集成和二次开发使用。

## 目录结构

```
scripts/
├── docx_to_config.py          # DOCX → JSON 主入口
├── generate_slides.py         # JSON → PPT 主入口
├── llm_client.py              # LLM 抽象层（多模型支持）
├── export_template_structure.py  # 模板结构导出工具
│
├── docx_processing/           # DOCX 解析子包
│   ├── __init__.py            # 统一导出
│   ├── constants.py           # 常量和正则表达式
│   ├── docx_parser.py         # 文档解析、图片提取
│   ├── docx_table_parser.py   # 表格解析
│   ├── template_utils.py      # 模板定义加载
│   ├── llm_processor.py       # LLM 调用逻辑
│   ├── llm_prompts.py         # Prompt 构建
│   ├── slide_filler.py        # 幻灯片内容填充
│   ├── special_pages.py       # 特殊页面（封面/目录/结束页）
│   └── json_utils.py          # JSON 解析工具
│
└── ppt_processing/            # PPT 渲染子包
    ├── __init__.py            # 统一导出
    ├── constants.py           # 常量定义
    ├── slide_builder.py       # 幻灯片构建主逻辑
    ├── shape_utils.py         # 形状处理
    ├── text_utils.py          # 文本填充
    ├── image_utils.py         # 图片处理
    ├── layout_utils.py        # 布局计算
    ├── xml_utils.py           # XML 操作
    └── connector_utils.py     # 连接线处理
```

## 数据流

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│  DOCX 讲稿  │ ──▶ │ generate_config  │ ──▶ │ JSON 配置   │
└─────────────┘     │   (LLM 规划)     │     └─────────────┘
                    └──────────────────┘            │
                                                    ▼
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│  最终 PPT   │ ◀── │  render_slides   │ ◀── │ 模板 PPTX   │
└─────────────┘     └──────────────────┘     └─────────────┘
```

---

## 快速集成

### 1. 完整流程：DOCX → PPT

```python
from pathlib import Path
from scripts.docx_to_config import generate_config_data
from scripts.generate_slides import render_slides

# 步骤 1: DOCX → JSON
config = generate_config_data(
    docx_path="讲稿.docx",
    template_json="template/工科模板1/template.json",
    template_list=None,              # 已废弃，传 None
    use_llm=True,                    # 启用 LLM 智能规划
    llm_provider="deepseek",         # deepseek/glm/taichu/local
    llm_model="deepseek-chat",
    llm_base_url=None,               # 使用默认地址
    metadata_overrides={             # 可选：覆盖元数据
        "course": "传感器技术",
        "college": "机械工程学院",
        "lecturer": "张三",
    },
    run_dir=Path("output"),          # 输出目录
    user_prompt=None,                # 可选：用户自定义提示
)

# 步骤 2: JSON → PPT
result = render_slides(
    template_path=Path("template/工科模板1/template.pptx"),
    config=config,
    output_name="slides.pptx",
    run_dir=Path("output"),
)

print(f"PPT 生成完成: {result['output_path']}")
print(f"共 {result['slides']} 页")
```

### 2. 仅解析 DOCX（不生成 PPT）

```python
from pathlib import Path
from scripts.docx_processing import parse_docx_blocks, load_template_defs

# 解析 DOCX，提取文本和图片
image_dir = Path("output/images")
blocks, has_marker, metadata = parse_docx_blocks("讲稿.docx", image_dir)

# blocks: 内容块列表，每个包含 text, images, template_hint
# has_marker: 是否包含【PPT1】等标记
# metadata: 提取的元数据（课程名、学院、主讲人）

for block in blocks:
    print(f"模板提示: {block.get('template_hint')}")
    print(f"文本: {block.get('text')[:100]}...")
    print(f"图片: {block.get('images')}")
```

### 3. 仅渲染 PPT（已有 JSON）

```python
import json
from pathlib import Path
from scripts.generate_slides import render_slides

# 加载已有的 JSON 配置
with open("config.json", encoding="utf-8") as f:
    config = json.load(f)

result = render_slides(
    template_path=Path("template/工科模板1/template.pptx"),
    config=config,
    output_name="slides.pptx",
    run_dir=Path("output"),
)
```

### 4. 导出模板结构

```python
from pathlib import Path
from scripts.export_template_structure import export_template_structure, ai_enrich_template

# 分析 PPT 模板，导出结构定义
template_data = export_template_structure(
    template_path=Path("template/template.pptx"),
    mode="semantic",      # semantic（语义分组）或 text（纯文本）
    include_pages=None,   # None 表示所有页，或 [1, 2, 4] 指定页码
)

# 可选：使用 AI 自动填充 hint、required、max_chars
enriched_data = ai_enrich_template(
    template_data=template_data,
    llm_provider="deepseek",
    llm_model="deepseek-chat",
)

# 保存为 template.json
import json
with open("template.json", "w", encoding="utf-8") as f:
    json.dump(enriched_data, f, ensure_ascii=False, indent=2)
```

---

## 核心 API 详解

### `generate_config_data()`

将 DOCX 讲稿转换为 JSON 配置，是整个流程的核心函数。

```python
def generate_config_data(
    docx_path: str,                  # DOCX 文件路径
    template_json: str,              # 模板定义 JSON 路径
    template_list: str,              # 已废弃，传 None
    use_llm: bool,                   # 是否启用 LLM
    llm_provider: str,               # LLM 提供商
    llm_model: Optional[str],        # LLM 模型名称
    llm_base_url: Optional[str],     # LLM API 地址（可选）
    metadata_overrides: Optional[Dict[str, str]],  # 元数据覆盖
    run_dir: Path,                   # 输出目录
    user_prompt: Optional[str] = None,  # 用户自定义提示
) -> Dict:
    """
    返回值:
        {
            "ppt_pages": [
                {
                    "page_type": "封面",
                    "template_page_num": 1,
                    "content": {...}
                },
                ...
            ]
        }
    """
```

**处理流程：**
1. 解析 DOCX，提取文本块和图片
2. 加载模板定义和全局配置
3. 调用 LLM 进行内容规划和分页
4. 填充每页内容
5. 添加特殊页面（封面、目录、结束页）

### `render_slides()`

根据 JSON 配置渲染 PPT。

```python
def render_slides(
    template_path: Path,    # 模板 PPTX 路径
    config: dict,           # JSON 配置字典
    output_name: str,       # 输出文件名
    run_dir: Path = None,   # 输出目录（可选）
) -> dict:
    """
    返回值:
        {
            "output_path": Path,  # 生成的 PPT 路径
            "run_dir": Path,      # 运行目录
            "slides": int,        # 幻灯片数量
        }
    """
```

### `export_template_structure()`

导出 PPT 模板的结构定义，用于生成 template.json。

```python
def export_template_structure(
    template_path: Path,              # 模板 PPTX 路径
    mode: str = "semantic",           # 导出模式
    include_pages: list[int] = None,  # 包含的页码
) -> dict:
    """
    返回值:
        {
            "template_prompt": {...},   # 全局配置
            "special_pages": {...},     # 特殊页面配置
            "manifest": [...],          # 页面清单
            "ppt_pages": [...]          # 页面详细定义
        }
    """
```

---

## LLM 客户端

### 支持的提供商

| 提供商 | 类名 | 环境变量 | 多模态 |
|--------|------|----------|--------|
| DeepSeek | `DeepSeekLLM` | `DEEPSEEK_API_KEY` | ❌ |
| 智谱 AI | `GLMLLM` | `GLM_API_KEY` | ✅ |
| 紫东太初 | `TaichuLLM` | `TAICHU_API_KEY` | ✅ |
| 本地部署 | `LocalLLM` | `LOCAL_LLM_BASE_URL` | - |
| Qwen vLLM | `QwenVLLM` | - | - |

### 直接使用 LLM 客户端

```python
from scripts.llm_client import DeepSeekLLM, GLMLLM, LocalLLM

# DeepSeek
llm = DeepSeekLLM(model="deepseek-chat")
response = llm.generate([
    {"role": "user", "content": "你好"}
])

# 智谱 GLM（多模态）
llm = GLMLLM(model="glm-4.5v")
response = llm.generate([
    {"role": "user", "content": [
        {"type": "text", "text": "描述这张图片"},
        {"type": "image_url", "image_url": {"url": "data:image/png;base64,..."}}
    ]}
])

# 本地部署（OpenAI 兼容接口）
llm = LocalLLM(
    base_url="http://localhost:8000/v1",
    model="qwen2-7b",
)
```

### 使用 `choose_llm()` 工厂函数

```python
from scripts.docx_processing import choose_llm

# 根据参数自动选择 LLM 实现
llm = choose_llm(
    use_llm=True,
    llm_provider="deepseek",
    llm_model="deepseek-chat",
    llm_base_url=None,
)

if llm:
    response = llm.generate([{"role": "user", "content": "你好"}])
```

---

## JSON 配置格式

### 完整示例

```json
{
  "ppt_pages": [
    {
      "page_type": "封面",
      "template_page_num": 1,
      "content": {
        "封面区": {
          "课程名称": "传感器技术",
          "主讲人": "张三",
          "学院": "机械工程学院"
        }
      }
    },
    {
      "page_type": "目录页",
      "template_page_num": 2,
      "content": {
        "目录区": {
          "目录项1": "第一章 概述",
          "目录项2": "第二章 温度传感器",
          "目录项3": "第三章 压力传感器"
        }
      }
    },
    {
      "page_type": "章节页",
      "template_page_num": 3,
      "content": {
        "章节区": {
          "一级标题": "第一章",
          "二级标题": "概述"
        }
      }
    },
    {
      "page_type": "图文页",
      "template_page_num": 8,
      "content": {
        "内容区": {
          "标题": "温度传感器原理",
          "要点1": "热电偶原理",
          "要点2": "热敏电阻原理",
          "图片": "/path/to/image.png"
        }
      }
    }
  ]
}
```

### 图片字段格式

图片字段支持两种格式：

```json
// 格式 1: 直接路径字符串
"图片": "/path/to/image.png"

// 格式 2: 对象格式（兼容旧版）
"图片": {
  "type": "image",
  "value": "/path/to/image.png"
}
```

---

## 模板配置 (template.json)

### 结构说明

```json
{
  "template_prompt": {
    "preprocess_guide": "LLM 预处理指导（可选）",
    "fill_guide": "LLM 填充指导（可选）",
    "section_tracking": true,
    "section_field_mappings": {
      "chapter": "'章节'、'一级标题'",
      "section": "'知识点'、'二级标题'"
    }
  },
  "special_pages": {
    "cover": 1,
    "toc": 2,
    "end": 28
  },
  "manifest": [
    {"template_page_num": 1, "page_type": "封面", "text_slots": 3, "image_slots": 0},
    {"template_page_num": 3, "page_type": "章节页", "text_slots": 2, "image_slots": 0}
  ],
  "ppt_pages": [
    {
      "page_type": "封面",
      "template_page_num": 1,
      "content": {
        "封面区": {
          "课程名称": {"type": "text", "hint": "填写课程名称", "required": true, "max_chars": 20, "value": ""},
          "主讲人": {"type": "text", "hint": "填写主讲人姓名", "required": true, "max_chars": 10, "value": ""}
        }
      },
      "meta": {
        "notes": "封面页，用于展示课程基本信息"
      }
    }
  ]
}
```

### 全局配置说明

| 字段 | 说明 |
|------|------|
| `template_prompt.preprocess_guide` | LLM 预处理时的额外指导 |
| `template_prompt.fill_guide` | LLM 填充内容时的额外指导 |
| `template_prompt.section_tracking` | 是否启用章节追踪（默认 true） |
| `template_prompt.section_field_mappings` | 章节字段映射关系 |
| `special_pages.cover` | 封面页模板编号（null 表示不添加） |
| `special_pages.toc` | 目录页模板编号（null 表示不添加） |
| `special_pages.end` | 结束页模板编号（null 表示不添加） |

---

## 子模块详解

### docx_processing 子包

| 模块 | 职责 |
|------|------|
| `constants.py` | 正则表达式、常量定义 |
| `docx_parser.py` | DOCX 解析、图片提取、章节追踪 |
| `docx_table_parser.py` | 表格内容提取 |
| `template_utils.py` | 加载和处理 template.json |
| `llm_processor.py` | LLM 调用封装（预处理、填充） |
| `llm_prompts.py` | Prompt 模板构建 |
| `slide_filler.py` | 内容填充逻辑 |
| `special_pages.py` | 封面/目录/结束页处理 |

### ppt_processing 子包

| 模块 | 职责 |
|------|------|
| `constants.py` | XML 命名空间、常量 |
| `slide_builder.py` | 幻灯片构建主逻辑 |
| `shape_utils.py` | 形状查找、删除、复制 |
| `text_utils.py` | 文本框填充、格式处理 |
| `image_utils.py` | 图片插入、缩放 |
| `layout_utils.py` | 布局计算、位置调整 |
| `xml_utils.py` | XML 操作、关系清理 |

---

## 错误处理

```python
from scripts.docx_to_config import generate_config_data

try:
    config = generate_config_data(...)
except FileNotFoundError as e:
    print(f"文件不存在: {e}")
except ValueError as e:
    print(f"配置错误: {e}")
except RuntimeError as e:
    print(f"LLM 调用失败: {e}")
```

常见错误：
- `FileNotFoundError`: DOCX 或模板文件不存在
- `ValueError`: 模板配置错误、未生成任何内容
- `RuntimeError`: LLM API 调用失败

---

## CLI 命令

```bash
# 完整流程
python scripts/docx_to_config.py \
    --docx 讲稿.docx \
    --template-json template/工科模板1/template.json \
    --use-llm \
    --llm-provider deepseek \
    --run-dir output

python scripts/generate_slides.py \
    --template template/工科模板1/template.pptx \
    --json output/config.json \
    --output slides.pptx

# 从预处理讲稿继续（跳过 LLM 预处理）
python scripts/docx_to_config.py \
    --from-preprocessed output/preprocessed_script.md \
    --template-json template/工科模板1/template.json \
    --template-pptx template/工科模板1/template.pptx \
    --use-llm

# 导出模板结构
python scripts/export_template_structure.py \
    --template template/template.pptx \
    --output template/template.json \
    --mode semantic \
    --ai-enrich
```

---

## 依赖

```
python-pptx>=0.6.21
python-docx>=0.8.11
Pillow>=9.0.0
lxml>=4.9.0
requests>=2.28.0
```
