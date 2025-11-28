# S2S（Script to Slides）

> 🚀 将讲稿自动转换为精美 PPT 的智能工具

[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)
[![Django](https://img.shields.io/badge/Django-5.2-green.svg)](https://djangoproject.com)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

S2S 是一个自动化 PPT 生成工具，支持从 DOCX 讲稿自动生成结构化 PPT，集成多种大模型（LLM）进行智能内容规划，并提供可视化的模板编辑器。

## ✨ 核心功能

- 📝 **讲稿转 PPT**：从 DOCX 讲稿自动生成 PPT，支持文字和图片
- 🤖 **AI 智能规划**：支持 DeepSeek、智谱 GLM、紫东太初等多种大模型
- 🎨 **模板制作向导**：可视化四步流程，轻松制作 PPT 模板
- 🔧 **模板编辑器**：在线编辑 PPT 元素，AI 一键命名，批量操作
- 📊 **配置模板管理**：JSON 配置可视化编辑，AI 自动填充提示信息
- 🌐 **Web 界面**：基于 Django 的现代化 Web 前端

## 🚀 快速开始

### 安装

```bash
# 克隆项目
git clone https://github.com/bonaparte233/S2S.git
cd S2S

# 创建虚拟环境
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 初始化数据库和用户
cd web
python manage.py migrate
python manage.py init_users
cd ..
```

### 启动服务

```bash
# macOS/Linux
./start_web.sh

# Windows
start_web.bat
```

访问 `http://127.0.0.1:8000/`

### 默认账户

| 用户名 | 密码 | 角色 |
|--------|------|------|
| `admin` | `admin123` | 管理员（全部权限） |
| `developer` | `dev123` | 开发者（模板编辑、LLM 配置） |
| `user` | `user123` | 普通用户（PPT 生成） |

## 📖 使用指南

### Web 界面（推荐）

1. **生成 PPT**：上传 DOCX 讲稿 → 选择模板 → 启用 AI → 生成
2. **模板制作向导**：上传 PPT → 元素命名 → 配置模板 → 发布
3. **开发者工具**：模板编辑器、配置编辑器、AI 批量命名

### 命令行

```bash
# 一键生成
python main.py --docx 讲稿.docx --use-llm --ppt-output output.pptx

# 仅生成 JSON 配置
python scripts/docx_to_config.py --docx 讲稿.docx --use-llm --output config.json

# 从 JSON 生成 PPT
python scripts/generate_slides.py --template template.pptx --json config.json --output slides.pptx
```

## 🤖 LLM 配置

### 支持的模型

| 提供商 | 模型 | 多模态 | 环境变量 |
|--------|------|--------|----------|
| DeepSeek | `deepseek-chat` | ❌ | `DEEPSEEK_API_KEY` |
| 智谱 AI | `glm-4.5v` | ✅ | `GLM_API_KEY` |
| 紫东太初 | `taichu4_vl_32b` | ✅ | `TAICHU_API_KEY` |
| 本地部署 | 自定义 | - | `LOCAL_LLM_BASE_URL` |

### 配置方式

**方式一：管理员预设（推荐）**

1. 访问 Django Admin (`/admin/`)
2. 创建「全局 LLM 配置」
3. 用户在生成时选择预设配置

**方式二：临时自定义**

- 开发者在生成页面手动输入 API Key 和模型参数

## 📁 项目结构

```
S2S/
├── main.py                 # 主入口
├── scripts/                # 核心脚本
│   ├── docx_to_config.py   # DOCX → JSON
│   ├── generate_slides.py  # JSON → PPT
│   ├── llm_client.py       # LLM 抽象层
│   └── export_template_structure.py
├── web/                    # Django Web 应用
│   ├── ppt_generator/      # PPT 生成模块
│   ├── templates/          # HTML 模板
│   └── static/             # 静态资源
├── template/               # PPT 模板和配置
└── requirements.txt
```

## 🛠️ 开发者工具

### 模板制作向导

四步流程制作 PPT 模板：

1. **上传 PPT**：上传 PPTX 文件，自动解析页面和元素
2. **元素命名**：为每个元素设置语义化名称（支持 AI 一键命名）
3. **配置模板**：设置元素的提示信息、必填项、字数限制
4. **发布模板**：保存到模板库，可直接用于 PPT 生成

### PPT 模板编辑器功能

- 🖼️ 可视化预览，元素高亮标注
- ✏️ 在线编辑元素名称和属性
- 🤖 AI 一键命名（逐页处理）
- 📦 批量操作（多选、隐藏、显示）
- 💾 自动保存编辑进度

## 🔧 高级配置

### 环境变量

```bash
# LLM API Keys
export DEEPSEEK_API_KEY="your-key"
export GLM_API_KEY="your-key"
export TAICHU_API_KEY="your-key"

# Django 配置
export SECRET_KEY="your-secret-key"
export DEBUG=False
export ALLOWED_HOSTS="your-domain.com"
```

### 生产部署

```bash
cd web
python manage.py collectstatic --noinput
gunicorn web_frontend.wsgi:application --bind 0.0.0.0:8000 --workers 4
```

## 📋 技术栈

- **后端**：Python 3.11、Django 5.2、python-pptx、python-docx
- **前端**：HTML5、CSS3、Vanilla JavaScript
- **数据库**：SQLite（开发）/ PostgreSQL（生产）
- **AI**：OpenAI SDK（兼容多种 LLM API）

## ❓ 常见问题

<details>
<summary><b>启动失败：ModuleNotFoundError</b></summary>

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

</details>

<details>
<summary><b>API 密钥错误</b></summary>

检查 Django Admin 中的全局 LLM 配置，或在生成页面手动输入正确的 API Key。
</details>

<details>
<summary><b>模板文件找不到</b></summary>

确保 `template/` 目录下有对应的 `.pptx` 和 `.json` 文件，或在页面上传自定义模板。
</details>

## 📝 更新日志

### v2.4.0 (2025-11-28)

- ✨ **模板制作向导**
  - 新增四步向导流程：上传 PPT → 元素命名 → 配置模板 → 发布
  - 支持 iframe 嵌入编辑器，统一的向导体验
  - 会话保存和恢复，随时继续编辑
- ✨ **PPT 模板编辑器**
  - AI 一键命名全部页（支持向导模式）
  - 批量操作：多选、隐藏、显示元素
  - 元素属性编辑面板优化
- ✨ **编辑会话管理**
  - 自动保存编辑进度
  - 编辑记录列表，一键恢复
  - 支持 wizard/ppt/config 三种会话类型
- 🐛 **问题修复**
  - 修复向导会话恢复时步骤定位问题
  - 修复 PPT 文件上传后发布失败的问题
  - 修复嵌入模式下导航栏显示问题

### v2.3.0 (2025-11-24)

- ✨ 多模态 LLM 增强（GLM-4.5V 支持）
- 🔧 LLM 配置重构（预设/自定义模式）
- 🎨 前端样式优化

### v2.2.0 (2025-11-22)

- ✨ 多模态 LLM 支持（紫东太初、智谱 GLM）
- ✨ 图片智能理解和放置

### v2.1.0 (2025-11-20)

- ✨ 开发者工具页面
- ✨ 配置模板 AI 填充

### v2.0.0 (2025-11-19)

- ✨ Django Web 前端
- ✨ 用户认证和权限系统
- ✨ 全局 LLM 配置

### v1.0.0

- 🎉 初始版本（DOCX → JSON → PPT）

---

## 📄 许可证

本项目仅供学习和研究使用。


