# S2S（Script to Slides）

将讲稿 **DOCX** 自动转换为结构化 **JSON**，再基于 **PPT 模板** 批量生成最终 **PPTX 幻灯片** 的工具链。

---

## 功能概述

- 从讲稿 DOCX 中解析段落、图片和基础元信息（课程名称、学院名称、主讲教师等）。
- 根据模板配置自动规划每一页内容（可使用大模型 LLM，也可在文稿中显式标记模板）。
- 按照既定 PPT 模板 (`template/template.pptx`) 填充文本与图片，生成统一风格的课件/汇报 PPT。
- 提供命令行入口和函数级 API，方便集成到 GUI 或其他系统中。

---

## 核心流程

1. **DOCX → 中间结构**：解析 DOCX，按段落拆分内容、提取图片，识别模板标记与元信息。
2. **生成 JSON 配置**：调用 `docx_to_config.generate_config_data`，生成形如 `{"ppt_pages": [...]}` 的配置文件。
3. **JSON → PPT**：调用 `generate_slides.render_slides`，复制 PPT 模板页并填充内容，输出最终 PPTX。

---

## 目录结构（关键文件）

- `main.py`：主入口，一条命令完成 DOCX → JSON → PPT。
- `docx_to_config.py`：DOCX → JSON 的核心逻辑与独立 CLI。
- `generate_slides.py`：JSON → PPT 的核心逻辑与独立 CLI。
- `llm_client.py`：大模型抽象与 DeepSeek / 本地模型 Provider 封装。
- `template/`：PPT 模板文件、模板定义 JSON、模板编号白名单以及示例讲稿 DOCX。
- `temp/`：运行目录，每次运行会在其中创建带时间戳和随机后缀的子目录, 包含中间 JSON、生成的 PPT 以及提取的图片。
- `archive/`：历史遗留文件。
- `requirements.txt`：Python 依赖列表。

---

## 安装

1. 准备好 Python 3 环境。
2. 在项目根目录执行：

```bash
pip install -r requirements.txt
```

---

## 快速开始（推荐）

在项目根目录运行：

```bash
python main.py \
  --docx path/to/讲稿.docx \
  --use-llm \
  --ppt-output output/我的课件.pptx
```

运行后：

- 会在 `temp/run-时间戳-随机值` 目录下生成：
  - `config.json`：中间 JSON 配置（只包含 `ppt_pages` 列表）。
  - `slides.pptx`：根据模板填充后的 PPT。
  - `images/`：从 DOCX 中提取并用于填充的图片（如有）。
- 同时会将最终 PPT 复制到 `--ppt-output` 指定路径。

### `main.py` 常用参数

- `--docx`（必选）：讲稿 DOCX 路径。
- `--template-json`（默认 `template/template.json`）：模板定义 JSON。
- `--template-list`（默认 `template/template.txt`）：允许使用的模板编号列表。
- `--template-ppt`（默认 `template/template.pptx`）：PPT 模板文件。
- `--run-dir`：自定义运行目录（默认在 `temp/` 下自动创建 `run-...` 目录）。
- `--config-name`（默认 `config.json`）：运行目录中的 JSON 文件名。
- `--slides-name`（默认 `slides.pptx`）：运行目录中的 PPT 文件名。
- `--ppt-output`：如需额外复制 PPT，请提供完整输出路径。
- `--use-llm`：是否启用大模型进行内容生成/排版。
- `--llm-provider`（默认 `deepseek`）：大模型提供商标识，当前代码支持 `deepseek` 和 `local`。
- `--llm-model`（默认 `deepseek-chat`）：模型名称，传给对应 Provider。
- `--course-name` / `--college-name` / `--lecturer-name`：覆盖课程名称、学院名称、主讲教师姓名。

---

## 分步使用

### 1. 只生成 JSON（DOCX → JSON）

```bash
python docx_to_config.py \
  --docx path/to/讲稿.docx \
  --use-llm \
  --output config/my_config.json
```

主要参数含义与 `main.py` 中同名参数一致：

- `--docx`：讲稿 DOCX 路径（必选）。
- `--template-json` / `--template-list`：模板定义与模板编号白名单。
- `--use-llm` / `--llm-provider` / `--llm-model`：控制是否启用 LLM 以及使用哪种 Provider。
- `--course-name` / `--college-name` / `--lecturer-name`：覆盖元信息。
- `--run-dir` / `--config-name`：控制运行目录与内部 JSON 文件名。
- `--output`：如需额外复制一份 JSON 到指定路径，提供完整路径；否则仅在运行目录中生成。

生成的 JSON 顶层结构为：`{"ppt_pages": [...]}`。

### 2. 只根据 JSON 生成 PPT（JSON → PPT）

```bash
python generate_slides.py \
  --template template/template.pptx \
  --json config/my_config.json \
  --output output/my_slides.pptx
```

- `--template`（必选）：PPTX 模板路径。
- `--json`（必选）：描述内容的 JSON 文件路径（必须包含 `ppt_pages` 列表）。
- `--output`（默认 `final_output.pptx`）：输出 PPT 文件名或完整路径。
- `--run-dir`：指定运行目录；若为绝对路径输出，脚本会在生成后复制到该路径。

---

## 大模型（LLM）配置

大模型相关逻辑定义在 `llm_client.py` 与 `docx_to_config.py` 中：

- 当传入 `use_llm=True` 时，`generate_config_data` 会通过 `choose_llm` 选择具体 LLM：
  - `provider = "deepseek"` 时使用 `DeepSeekLLM`；
  - `provider = "local"` 时使用 `LocalLLM`；
  - 其他值会抛出错误：`暂不支持的大模型提供商：{provider}`。
- 如果讲稿中**没有任何模板标记**且未启用 LLM，则会抛出：
  - `讲稿未指定 PPT 标记且未启用 LLM，无法自动分配模板。`

### DeepSeek 模型

- 需要环境变量：
  - `DEEPSEEK_API_KEY`（必需）。
  - `DEEPSEEK_BASE_URL`（可选，默认 `https://api.deepseek.com`）。
- 默认模型名称为 `deepseek-chat`，也可通过命令行 `--llm-model` 覆盖。

### 本地/自建模型

- 使用环境变量：
  - `LOCAL_LLM_BASE_URL`（可选，默认 `http://127.0.0.1:8000/v1`）。
  - `LOCAL_LLM_MODEL`（可选，默认 `local-model`）。
  - `LOCAL_LLM_API_KEY`（可选，本地服务需要鉴权时使用）。

---

## 模板与讲稿约定

### 模板配置

- `template/template.json` 中的 `manifest` 定义了各模板页：
  - `template_page_num`：在 PPT 模板中的页号。
  - `page_type`：页面类型（如“封面页”“章节页”“图文页”等）。
  - `text_slots` / `image_slots`：文本/图片槽位数量。
- 后续字段还会为每种页面类型定义具体的文本/图片字段及其路径，并提供布局、使用场景、风格、备注等元信息，作为 LLM 提示的一部分。
- `template/template.txt` 中列出了允许使用的模板编号，若讲稿中指定的编号不在其中，会报错提示模板未定义或未被允许。

### 讲稿书写规范与模板标记

- 代码中使用的模板标记正则为：`【PPTn】`，例如：
  - `【PPT1】`、`【PPT2】`、`【PPT18】` 等。
- 在解析 DOCX 时，脚本会：
  - 按段落扫描这些标记，将标记及前后文本划分为块（block），并记录 `template_hint`。
  - 同时识别形如：
    - `课程名称：XXX`
    - `学院名称：YYY`
    - `主讲教师：ZZZ`
    并写入内部 `metadata["course" | "college" | "lecturer"]`。
- 在填充模板时，`metadata` 会通过 `_apply_metadata_overrides` 自动套入如课程名、学院名、主讲人等字段。

> 提示：
>
> - 不启用 LLM 时，建议在讲稿中为每个需要生成 PPT 的内容块添加 `【PPTn】` 标记；
> - 如不想在文稿中写标记，则需要启用 LLM，让模型自动规划 PPT 结构。

---

## 输出与运行目录

- 脚本会在 `temp/` 下创建带时间戳和随机后缀的运行目录，例如：
  - `run-YYYYMMDD-HHMMSS-xxxx`（`main.py`）。
  - `script-YYYYMMDD-HHMMSS-xxxx`（直接使用 `docx_to_config.py` 时）。
- 每个运行目录一般包含：
  - `config.json`：当前讲稿对应的 PPT 配置 JSON。
  - `slides.pptx` 或其他命名的 PPT 文件。
  - `images/`：从 DOCX 中提取并用于填充的图片（如有）。

---

## 示例文件

- `template/模板.docx`、`template/天文模板.docx`：示例讲稿。
- `output/` 下的多份 PPTX（如 `output/output.pptx`、`output/final.pptx` 等）：示例生成结果。
- `images/` 下的 PNG/JPG 图片：示例图片资源，可在模板或讲稿中引用。

---

## 二次开发与集成

在其他程序中集成本项目时，可以直接调用 Python 函数而不是通过命令行：

- 从 DOCX 生成 JSON：
  - 调用 `docx_to_config.generate_config_data(docx_path, template_json, template_list, use_llm, llm_provider, llm_model, metadata_overrides, run_dir)`。
- 从 JSON 生成 PPT：
  - 调用 `generate_slides.render_slides(template_path, config, output_name, run_dir=None)`。

`main.py` 仅仅是对上述两步的封装，方便命令行和 GUI 共同复用。
