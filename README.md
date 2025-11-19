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

### 全局 LLM 配置

管理员可以在 Django Admin 后台配置全局的 LLM 设置，这样普通用户无需输入 API 密钥即可使用大模型功能。

**配置步骤：**

1. 访问 `http://127.0.0.1:8000/admin/`
2. 使用 `admin` / `admin123` 登录
3. 点击"全局LLM配置"
4. 配置：
   - LLM 供应商（DeepSeek / 本地部署 / 自定义服务）
   - LLM 模型（如 `deepseek-chat`）
   - API Key（全局默认密钥）
   - 服务器地址（可选）
   - 默认系统 Prompt（可选）
5. 保存

**配置优先级：**

- 开发者/管理员在生成页面输入的配置会临时覆盖全局配置
- 普通用户勾选"使用大模型"后自动使用全局配置

### Web 功能特性

- ✅ **用户认证**：强制登录，按角色分配权限
- ✅ **在线上传**：支持上传 DOCX 讲稿和自定义 PPT 模板
- ✅ **LLM 配置**：支持 DeepSeek、本地部署、自定义服务
- ✅ **实时状态**：自动轮询生成状态，完成后自动刷新
- ✅ **历史记录**：查看个人生成历史，下载 PPT 和 JSON
- ✅ **全局配置**：管理员统一管理 API 密钥
- ✅ **模板导出**：开发者可导出模板结构为 JSON（语义/文本模式）

---

## 核心流程

1. **DOCX → 中间结构**：解析 DOCX，按段落拆分内容、提取图片，识别模板标记与元信息。
2. **生成 JSON 配置**：调用 `docx_to_config.generate_config_data`，生成形如 `{"ppt_pages": [...]}` 的配置文件。
3. **JSON → PPT**：调用 `generate_slides.render_slides`，复制 PPT 模板页并填充内容，输出最终 PPTX。

---

## 目录结构（关键文件）

### 核心脚本

- `main.py`：主入口，一条命令完成 DOCX → JSON → PPT。
- `scripts/docx_to_config.py`：DOCX → JSON 的核心逻辑与独立 CLI。
- `scripts/generate_slides.py`：JSON → PPT 的核心逻辑与独立 CLI。
- `scripts/llm_client.py`：大模型抽象与 DeepSeek / 本地模型 Provider 封装。
- `scripts/export_template_structure.py`：导出模板结构为 JSON（语义/文本模式）。

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

### v2.0.0

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
