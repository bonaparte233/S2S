# S2S（Script to Slides）

将讲稿 **DOCX** 自动转换为结构化 **JSON**，再基于 **PPT 模板** 批量生成最终 **PPTX 幻灯片** 的工具链。

---

## 功能概述

- 从讲稿 DOCX 中解析段落、图片和基础元信息（课程名称、学院名称、主讲教师等）。
- 根据模板配置自动规划每一页内容（可使用大模型 LLM，也可在文稿中显式标记模板）。
- 按照既定 PPT 模板 (`template/template.pptx`) 填充文本与图片，生成统一风格的课件/汇报 PPT。
- 提供命令行入口、函数级 API 和 **Web 前端界面**，方便不同场景使用。

---

## 🌐 Web 前端（推荐）

项目提供了基于 Django 的 Web 前端，支持在线上传讲稿、配置模板、启用大模型，并在线查看生成历史。

### 快速启动

在项目根目录运行：

```bash
./start_web.sh
```

然后访问：`http://127.0.0.1:8000/`

### 用户系统

系统内置三种用户角色：

| 用户名 | 密码 | 角色 | 权限 |
|--------|------|------|------|
| `admin` | `admin123` | 管理员 | 所有权限 + Django Admin 后台访问 |
| `developer` | `dev123` | 开发者 | LLM 配置、JSON 下载、模板导出 |
| `user` | `user123` | 普通用户 | 基础 PPT 生成功能 |

**初始化用户：**

```bash
source .venv/bin/activate
cd web
python manage.py init_users
```

### LLM 配置管理

系统支持两种 LLM 配置方式：

#### 1. 预设配置（推荐）

管理员可以在 Django Admin 后台创建多个预设配置，用户直接选择使用：

**配置步骤：**

1. 访问 `http://127.0.0.1:8000/admin/`
2. 使用 `admin` / `admin123` 登录
3. 点击"全局LLM配置"
4. 创建新配置：
   - 配置名称（如"DeepSeek 默认配置"、"GLM 多模态配置"）
   - LLM 供应商（DeepSeek / 紫东太初多模态模型 / 智谱AI (GLM) / 本地部署 / 自定义服务）
   - LLM 模型（如 `deepseek-chat`、`glm-4v-plus`、`taichu4_vl_32b`）
   - API Key（全局默认密钥）
   - 服务器地址（可选）
   - 默认系统 Prompt（可选）
   - 是否设为默认配置
5. 保存

**使用方式：**

- 用户在生成页面勾选"使用大模型"
- 选择"使用预设配置"
- 从下拉框选择管理员创建的配置
- 无需填写 API Key 等敏感信息

#### 2. 自定义配置

开发者和管理员可以临时使用自定义配置：

- 选择"自定义配置"
- 手动填写 LLM 供应商、模型名称、API Key 等
- 仅在当前生成任务中使用

### Web 功能特性

- ✅ **用户认证**：强制登录，按角色分配权限
- ✅ **在线上传**：支持上传 DOCX 讲稿和自定义 PPT 模板
- ✅ **配置模板管理**：支持上传和选择 JSON 配置模板
- ✅ **LLM 配置**：支持 DeepSeek、本地部署、自定义服务
- ✅ **实时状态**：自动轮询生成状态，完成后自动刷新
- ✅ **历史记录**：查看个人生成历史，下载 PPT 和 JSON
- ✅ **全局配置**：管理员统一管理 API 密钥
- ✅ **开发者工具**：
  - 从 PPTX 生成配置模板（语义/文本模式）
  - 在线编辑配置模板（hint、required、max_chars、notes）
  - AI 一键填充配置模板（使用 LLM 自动生成提示信息）

---

## 核心流程

1. **DOCX → 中间结构**：解析 DOCX，按段落拆分内容、提取图片，识别模板标记与元信息。
2. **生成 JSON 配置**：调用 `docx_to_config.generate_config_data`，生成形如 `{"ppt_pages": [...]}` 的配置文件。
3. **JSON → PPT**：调用 `generate_slides.render_slides`，复制 PPT 模板页并填充内容，输出最终 PPTX。

---

## 目录结构（关键文件）

### 核心脚本

- `main.py`：主入口，一条命令完成 DOCX → JSON → PPT。
- `scripts/docx_to_config.py`：DOCX → JSON 的核心逻辑与独立 CLI，支持多模态图片处理。
- `scripts/generate_slides.py`：JSON → PPT 的核心逻辑与独立 CLI。
- `scripts/llm_client.py`：大模型抽象层，支持 DeepSeek、紫东太初多模态、智谱AI (GLM)、本地模型等多种 Provider。
- `scripts/export_template_structure.py`：导出模板结构为 JSON（语义/文本模式），支持 AI 填充配置。

### Web 前端

- `web/`：Django Web 应用
  - `ppt_generator/`：PPT 生成应用
    - `models.py`：数据模型（PPTGeneration、GlobalLLMConfig）
    - `views.py`：视图函数（上传、生成、历史记录等）
    - `forms.py`：表单定义
    - `admin.py`：Django Admin 配置
  - `templates/`：HTML 模板
  - `static/`：静态资源（CSS、JS、图片）
  - `media/`：用户上传的文件和生成的输出
- `start_web.sh`：Web 服务器启动脚本

### 资源文件

- `template/`：PPT 模板文件、模板定义 JSON、模板编号白名单以及示例讲稿 DOCX。
- `temp/`：运行目录，每次运行会在其中创建带时间戳和随机后缀的子目录, 包含中间 JSON、生成的 PPT 以及提取的图片。
- `archive/`：历史遗留文件。
- `requirements.txt`：Python 依赖列表。

---

## 安装

1. 准备好 Python 3.11 环境。
2. 在项目根目录执行：

```bash
# 创建虚拟环境
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 初始化 Web 数据库（如需使用 Web 前端）
cd web
python manage.py migrate
python manage.py init_users
cd ..
```

---

## 快速开始

### 方式一：Web 界面（推荐）

```bash
./start_web.sh
```

然后访问 `http://127.0.0.1:8000/`，使用以下账户登录：

- 管理员：`admin` / `admin123`
- 开发者：`developer` / `dev123`
- 普通用户：`user` / `user123`

### 方式二：命令行

在项目根目录运行：

```bash
# 激活虚拟环境
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 运行生成
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
  - `provider = "deepseek"` 时使用 `DeepSeekLLM`
  - `provider = "taichu"` 时使用 `TaichuLLM`（紫东太初多模态模型）
  - `provider = "glm"` 时使用 `GLMLLM`（智谱AI）
  - `provider = "local"` 时使用 `LocalLLM`
  - `provider = "qwen"` 时使用 `QwenVLLM`（通义千问 vLLM 部署）
  - 其他值会抛出错误：`暂不支持的大模型提供商：{provider}`
- 如果讲稿中**没有任何模板标记**且未启用 LLM，则会抛出：
  - `讲稿未指定 PPT 标记且未启用 LLM，无法自动分配模板。`

### 支持的 LLM 提供商

#### DeepSeek

- 需要环境变量：
  - `DEEPSEEK_API_KEY`（必需）
  - `DEEPSEEK_BASE_URL`（可选，默认 `https://api.deepseek.com`）
- 默认模型名称为 `deepseek-chat`，也可通过命令行 `--llm-model` 覆盖
- 支持纯文本生成

#### 紫东太初多模态模型 (Taichu)

- 需要环境变量：
  - `TAICHU_API_KEY`（必需）
  - `TAICHU_BASE_URL`（可选，默认 `https://platform.wair.ac.cn/maas/v1`）
- 默认模型名称为 `taichu4_vl_32b`
- **支持多模态**：可以处理包含图片的讲稿，图片会以 base64 格式发送给模型
- 适用于需要理解图片内容的场景

#### 智谱AI (GLM)

- 需要环境变量：
  - `GLM_API_KEY`（必需）
  - `GLM_BASE_URL`（可选，默认 `https://open.bigmodel.cn/api/paas/v4/`）
- 推荐模型：`glm-4v-plus`（多模态）、`glm-4.5v`（多模态）
- **支持多模态**：可以处理包含图片的讲稿
- 使用 OpenAI 兼容接口

#### 本地/自建模型

- 使用环境变量：
  - `LOCAL_LLM_BASE_URL`（可选，默认 `http://127.0.0.1:8000/v1`）
  - `LOCAL_LLM_MODEL`（可选，默认 `local-model`）
  - `LOCAL_LLM_API_KEY`（可选，本地服务需要鉴权时使用）
- 适用于自建的 OpenAI 兼容 API 服务

#### 通义千问 vLLM (Qwen)

- 需要环境变量：
  - `QWEN_BASE_URL`（必需）
  - `QWEN_MODEL`（可选，默认 `Qwen2-VL-7B-Instruct`）
  - `QWEN_API_KEY`（可选）
- 适用于 vLLM 部署的通义千问模型

### 多模态功能

当讲稿中包含图片时，系统会自动：

1. **提取图片**：从 DOCX 中提取所有图片并保存到 `images/` 目录
2. **图片编码**：将图片转换为 base64 格式
3. **多模态消息**：构建包含文本和图片的多模态消息
4. **智能理解**：
   - 模型会理解图片内容
   - 分析图片与上下文文本的关系（图片通常在相关文本下方）
   - 将图片放入合适的 PPT 页面
   - 根据图片内容优化文本描述

**支持多模态的模型**：

- 紫东太初多模态模型 (`taichu4_vl_32b`)
- 智谱AI GLM (`glm-4v-plus`, `glm-4.5v`)

**多模态 Prompt 增强**：

- 强调字数限制约束，防止生成内容超出 `max_chars` 限制
- 明确图片与文本的位置关系
- 强调图片的重要性，避免被忽略

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

## 开发者工具

### 导出模板结构

从 PPTX 模板导出配置模板 JSON：

```bash
python scripts/export_template_structure.py \
    --template template/template.pptx \
    --output template/exported_template.json \
    --mode semantic
```

**参数说明：**

- `--template`（必选）：PPTX 模板路径
- `--output`（必选）：导出 JSON 的输出路径
- `--mode`（默认 `semantic`）：导出模式
  - `semantic`：仅导出命名规范（含"xx区"等）的元素
  - `text`：导出所有可编辑文本框（忽略图片/背景）
- `--include`：可选，逗号分隔的页码列表，仅导出这些幻灯片，例如：`1,2,4`

### AI 填充配置模板

使用 AI 自动填充配置模板中的 `hint`、`required`、`max_chars` 和 `notes` 字段：

```bash
python scripts/export_template_structure.py \
    --template template/template.pptx \
    --output template/exported_template.json \
    --mode semantic \
    --ai-enrich \
    --llm-provider deepseek \
    --llm-model deepseek-chat
```

**AI 填充参数：**

- `--ai-enrich`：启用 AI 填充
- `--llm-provider`（默认 `deepseek`）：LLM 提供商（deepseek/local/qwen）
- `--llm-model`：LLM 模型名称
- `--llm-base-url`：LLM 服务器地址（仅在 local/qwen 时需要）

**注意：**

- AI 生成的内容可能不准确，需要人工审核和修改
- 需要设置环境变量 `DEEPSEEK_API_KEY`（如果使用 DeepSeek）
- 每个页面会调用一次 LLM，模板页面较多时可能需要较长时间

### Web 界面使用

开发者也可以在 Web 界面中使用这些功能：

1. 访问"开发者工具"页面（需要 `developer` 或 `admin` 账户）
2. **生成配置模板** Tab：
   - 上传 PPTX 模板文件
   - 选择导出模式（语义/文本）
   - 点击"生成配置模板"
   - 下载或直接编辑
3. **编辑配置模板** Tab：
   - 上传现有 JSON 配置文件
   - 点击"🤖 AI 一键填充"自动生成提示信息
   - 在左侧选择页面，在右侧编辑字段
   - 下载编辑后的配置模板

---

## 示例文件

- `template/模板.docx`、`template/天文模板.docx`：示例讲稿。
- `output/` 下的多份 PPTX（如 `output/output.pptx`、`output/final.pptx` 等）：示例生成结果。
- `images/` 下的 PNG/JPG 图片：示例图片资源，可在模板或讲稿中引用。

---

## 二次开发与集成

### Python API

在其他程序中集成本项目时，可以直接调用 Python 函数而不是通过命令行：

- 从 DOCX 生成 JSON：
  - 调用 `scripts.docx_to_config.generate_config_data(docx_path, template_json, template_list, use_llm, llm_provider, llm_model, metadata_overrides, run_dir)`。
- 从 JSON 生成 PPT：
  - 调用 `scripts.generate_slides.render_slides(template_path, config, output_name, run_dir=None)`。

`main.py` 仅仅是对上述两步的封装，方便命令行和 GUI 共同复用。

### Web 集成

Web 前端基于 Django 5.2.8 开发，可以作为独立服务部署：

**开发环境：**

```bash
./start_web.sh
```

**生产环境：**

```bash
source .venv/bin/activate
cd web

# 收集静态文件
python manage.py collectstatic --noinput

# 使用 Gunicorn 运行
gunicorn web_frontend.wsgi:application --bind 0.0.0.0:8000 --workers 4
```

**环境变量配置：**

在 `web/web_frontend/settings.py` 中可以配置：

- `SECRET_KEY`：Django 密钥
- `DEBUG`：调试模式
- `ALLOWED_HOSTS`：允许的主机名
- `MEDIA_ROOT`：媒体文件存储路径
- `S2S_TEMPLATE_DIR`：模板目录路径
- `S2S_TEMP_DIR`：临时文件目录路径

---

## 技术栈

### 后端

- **Python 3.11**
- **Django 5.2.8**：Web 框架
- **python-pptx**：PPT 文件操作
- **python-docx**：DOCX 文件解析
- **OpenAI SDK**：LLM 接口调用

### 前端

- **HTML5 + CSS3**
- **Vanilla JavaScript**：无框架依赖
- **Fetch API**：异步请求

### 数据库

- **SQLite**：开发环境默认数据库
- 支持 PostgreSQL、MySQL 等（生产环境推荐）

---

## 常见问题

### 1. Web 服务器启动失败

**问题：** `ModuleNotFoundError: No module named 'django'`

**解决：**

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. 生成失败：API 密钥错误

**问题：** `AuthenticationError: Invalid API key`

**解决：**

- 检查全局 LLM 配置中的 API 密钥是否正确
- 或在生成页面手动输入正确的 API 密钥

### 3. 模板文件找不到

**问题：** `FileNotFoundError: 模板文件不存在`

**解决：**

- 确保 `template/template.pptx` 存在
- 或上传自定义模板文件

### 4. 权限不足

**问题：** `403 Forbidden`

**解决：**

- 确认使用正确的用户角色登录
- 开发者功能需要 `developer` 或 `admin` 账户

---

## 许可证

本项目仅供学习和研究使用。

---

## 更新日志

### v2.3.0 (2025-11-24)

- ✨ **多模态 LLM 支持增强**
  - 修复 user_prompt 传递问题，确保用户自定义 prompt 在所有生成模式下生效
  - 增强字数限制约束，使用更强的语气和警告格式
  - 增强多模态指令，明确图片与文本的位置关系，强调图片重要性
  - 新增 GLM 多模态模型支持（`glm-4v-plus`, `glm-4.5v`）
  - 新增 `_is_multimodal_llm()` 辅助函数，统一多模态模型检测
- 🔧 **LLM 配置重构**
  - 将"LLM 供应商选择"改为"LLM 配置选择"
  - 支持预设配置和自定义配置两种模式
  - 预设配置模式：用户选择管理员创建的配置，无需填写 API Key
  - 自定义配置模式：开发者/管理员临时输入配置信息
  - 修复 Admin 界面显示问题，正确显示使用的 LLM 提供商
- 🎨 **前端优化**
  - 优化 LLM 配置方式区域布局，增加上边距
  - 统一单选按钮样式，使用 radio-group 和 radio-label 类
  - 修复 Django 5.2 模板语法错误（比较运算符两边必须有空格）
- 📝 更新文档

### v2.2.0 (2025-11-22)

- ✨ 新增多模态 LLM 支持
  - 支持紫东太初多模态模型 (Taichu-VL)
  - 支持智谱AI (GLM) 多模态模型
  - 自动提取 DOCX 中的图片并以 base64 格式发送给模型
  - 多模态消息构建（文本 + 图片）
- ✨ 新增多个 LLM Provider
  - 紫东太初多模态模型 (`TaichuLLM`)
  - 智谱AI (`GLMLLM`)
  - 通义千问 vLLM (`QwenVLLM`)
- 🔧 优化 LLM 配置管理
  - 支持多个全局 LLM 配置
  - 支持设置默认配置
  - 配置名称和描述字段
- 📝 更新文档

### v2.1.0 (2025-11-20)

- ✨ 新增配置模板管理功能
  - 支持上传和选择 JSON 配置模板
  - 配置模板与 PPTX 模板关联
  - 自动匹配配置模板
- ✨ 新增开发者工具页面
  - 从 PPTX 生成配置模板（语义/文本模式）
  - 在线编辑配置模板（左右分栏布局）
  - AI 一键填充配置模板（hint、required、max_chars、notes）
- ✨ 新增 `ai_enrich_template()` 函数
  - 使用 LLM 自动生成配置提示信息
  - 支持 CLI 和 Web 两种调用方式
- 🔧 优化配置模板 UI 设计
  - 移动到 LLM 配置区域
  - 改进交互逻辑
- 📝 更新文档

### v2.0.0 (2025-11-19)

- ✨ 新增 Django Web 前端
- ✨ 新增用户认证和权限系统
- ✨ 新增全局 LLM 配置功能
- ✨ 新增在线生成历史记录
- ✨ 新增实时状态轮询
- ✨ 新增模板导出功能（开发者）
- 🔧 优化代码结构，将核心脚本移至 `scripts/` 目录
- 🔧 改进 LLM 配置逻辑，支持配置优先级
- 📝 更新文档

### v1.0.0

- 🎉 初始版本
- ✅ 支持 DOCX → JSON → PPT 转换
- ✅ 支持 DeepSeek 和本地 LLM
- ✅ 支持命令行和函数调用
