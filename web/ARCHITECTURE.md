# S2S Web 应用架构

## 系统架构图

```
┌─────────────────────────────────────────────────────────────────┐
│                         用户浏览器                                │
│                    http://127.0.0.1:8000                        │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         │ HTTP Request/Response
                         │
┌────────────────────────▼────────────────────────────────────────┐
│                      Django Web 应用                             │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                    URL 路由层                             │  │
│  │  s2s_web/urls.py → ppt_generator/urls.py                │  │
│  └────────────────────┬─────────────────────────────────────┘  │
│                       │                                         │
│  ┌────────────────────▼─────────────────────────────────────┐  │
│  │                   视图层 (Views)                          │  │
│  │  - index()                  : 首页 + 上传表单            │  │
│  │  - generation_detail()      : 生成详情页                 │  │
│  │  - start_generation()       : 启动生成 (AJAX)           │  │
│  │  - check_status()           : 检查状态 (AJAX)           │  │
│  │  - download_ppt()           : 下载 PPT                  │  │
│  │  - download_json()          : 下载 JSON (开发者)        │  │
│  │  - history()                : 历史记录                   │  │
│  │  - developer_tools()        : 开发者工具 (开发者)       │  │
│  │  - generate_config_template(): 生成配置模板 (AJAX)      │  │
│  │  - ai_enrich_template_view(): AI 填充配置 (AJAX)        │  │
│  │  - login/logout()           : 用户认证                   │  │
│  └────────────────────┬─────────────────────────────────────┘  │
│                       │                                         │
│  ┌────────────────────▼─────────────────────────────────────┐  │
│  │                 表单层 (Forms)                            │  │
│  │  PPTGenerationForm : 文件上传 + 配置验证                 │  │
│  └────────────────────┬─────────────────────────────────────┘  │
│                       │                                         │
│  ┌────────────────────▼─────────────────────────────────────┐  │
│  │                 模型层 (Models)                           │  │
│  │  PPTGeneration:                                          │  │
│  │    - user           : 所属用户 (ForeignKey)              │  │
│  │    - docx_file      : 讲稿文件                           │  │
│  │    - template_file  : 模板文件                           │  │
│  │    - output_ppt     : 生成的 PPT                         │  │
│  │    - output_json    : 生成的 JSON                        │  │
│  │    - status         : 状态 (pending/processing/...)     │  │
│  │    - use_llm        : 是否使用 LLM                       │  │
│  │    - llm_provider   : LLM 供应商                         │  │
│  │    - llm_model      : LLM 模型                           │  │
│  │    - llm_api_key    : API 密钥 (可选)                    │  │
│  │    - metadata       : 课程/学院/讲师信息                  │  │
│  │                                                          │  │
│  │  GlobalLLMConfig (单例):                                 │  │
│  │    - llm_provider   : 全局 LLM 供应商                    │  │
│  │    - llm_model      : 全局 LLM 模型                      │  │
│  │    - llm_api_key    : 全局 API 密钥                      │  │
│  │    - llm_base_url   : 服务器地址                         │  │
│  │    - default_prompt : 默认系统 Prompt                    │  │
│  └────────────────────┬─────────────────────────────────────┘  │
│                       │                                         │
│  ┌────────────────────▼─────────────────────────────────────┐  │
│  │               数据库层 (SQLite)                           │  │
│  │  db.sqlite3 : 存储生成记录                               │  │
│  └──────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
                         │
                         │ 调用后端脚本
                         │
┌────────────────────────▼────────────────────────────────────────┐
│                    S2S 后端处理层                                │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  scripts/docx_to_config.py                               │  │
│  │  - parse_docx_blocks()    : 解析 DOCX                   │  │
│  │  - generate_config_data() : 生成 JSON 配置              │  │
│  └──────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  scripts/generate_slides.py                              │  │
│  │  - build_from_json()      : 复制模板页                  │  │
│  │  - render_slides()        : 填充内容生成 PPT            │  │
│  └──────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  scripts/llm_client.py (可选)                            │  │
│  │  - LLM 智能规划内容                                      │  │
│  └──────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  scripts/export_template_structure.py                    │  │
│  │  - export_template_structure() : 导出模板结构           │  │
│  │  - ai_enrich_template()        : AI 填充配置            │  │
│  └──────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
                         │
                         │ 读写文件
                         │
┌────────────────────────▼────────────────────────────────────────┐
│                      文件系统                                    │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  template/          : 模板文件                           │  │
│  │    - template.pptx  : PPT 模板                          │  │
│  │    - template.json  : 模板定义                          │  │
│  │    - template.txt   : 模板列表                          │  │
│  └──────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  web/media/         : 用户上传和生成文件                 │  │
│  │    - uploads/       : 上传的 DOCX/PPTX                  │  │
│  │    - outputs/       : 生成的 PPT                        │  │
│  │    - configs/       : 生成的 JSON                       │  │
│  └──────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  temp/              : 临时文件                           │  │
│  │    - web-{id}/      : 每次生成的临时目录                 │  │
│  │      - images/      : 提取的图片                         │  │
│  │      - config.json  : 中间配置                          │  │
│  │      - slides.pptx  : 临时 PPT                          │  │
│  └──────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

## 数据流程

### 1. 用户上传流程

```
用户选择文件
    ↓
表单验证 (forms.py)
    ↓
创建 PPTGeneration 记录 (status=pending)
    ↓
保存文件到 media/uploads/
    ↓
跳转到详情页
```

### 2. PPT 生成流程

```
用户点击"开始生成"
    ↓
AJAX 请求 start_generation()
    ↓
更新状态为 processing
    ↓
获取 LLM 配置
    ├─ 用户提供配置？→ 使用用户配置（临时覆盖）
    └─ 否 → 使用全局配置（GlobalLLMConfig）
    ↓
调用 generate_config_data()
    ├─ 解析 DOCX 文件
    ├─ 提取文本和图片
    ├─ (可选) 调用 LLM 规划
    └─ 生成 JSON 配置
    ↓
调用 render_slides()
    ├─ 复制模板页
    ├─ 填充文本内容
    ├─ 替换图片
    └─ 生成最终 PPT
    ↓
保存文件到 media/outputs/
    ↓
更新状态为 completed
    ↓
返回下载链接
```

### 3. 状态轮询流程

```
详情页加载
    ↓
检查状态
    ├─ pending → 显示"开始生成"按钮
    ├─ processing → 显示加载动画 + 每 2 秒轮询状态
    ├─ completed → 显示下载按钮
    └─ failed → 显示错误信息
```

## 权限系统

### 用户角色

| 角色 | 权限 | 说明 |
|------|------|------|
| **管理员** (admin) | 所有权限 + Django Admin | 超级用户，可访问后台管理 |
| **开发者** (developer) | LLM 配置、JSON 下载、模板导出 | 可配置 LLM 参数和导出模板 |
| **普通用户** (user) | 基础 PPT 生成 | 只能上传文件和生成 PPT |

### 权限控制

- **强制登录**: 所有页面需要 `@login_required` 装饰器
- **角色检查**: 使用 Django Groups 和 Permissions
- **视图级权限**: 使用 `@permission_required` 装饰器
- **模板级权限**: 使用 `{% if is_developer %}` 条件判断
- **记录隔离**: 用户只能查看自己的生成记录

### LLM 配置优先级

1. **开发者/管理员**: 可在生成页面输入配置（临时覆盖全局配置）
2. **普通用户**: 勾选"使用大模型"后自动使用全局配置
3. **全局配置**: 管理员在 Django Admin 中配置，对所有用户生效

## 技术栈

### 前端

- **HTML5**: 语义化标记
- **CSS3**: 扁平化设计，响应式布局
- **JavaScript**: AJAX 请求，状态轮询，表单验证

### 后端

- **Django 5.2.8**: Web 框架
- **SQLite**: 数据库（开发环境）
- **python-pptx**: PPT 处理
- **python-docx**: DOCX 处理
- **Pillow**: 图片处理
- **OpenAI SDK**: LLM 接口调用

### 文件处理

- **FileField**: Django 文件上传
- **ContentFile**: 文件内容处理
- **Path**: 路径操作

### 认证与权限

- **Django Auth**: 用户认证系统
- **Groups & Permissions**: 角色和权限管理
- **Context Processors**: 自定义上下文处理器

## 安全考虑

1. **CSRF 保护**: Django 内置 CSRF token
2. **文件验证**: 检查文件类型和大小
3. **路径安全**: 使用 Path 对象，避免路径遍历
4. **错误处理**: 捕获异常，避免敏感信息泄露
5. **用户认证**: 强制登录，所有页面需要认证
6. **权限控制**: 基于角色的访问控制（RBAC）
7. **记录隔离**: 用户只能访问自己的记录
8. **API 密钥保护**: 全局配置中的密钥仅管理员可见

## 性能优化

1. **异步处理**: 使用 AJAX 避免页面阻塞
2. **状态轮询**: 仅在 processing 状态时轮询
3. **文件缓存**: 生成的文件保存在 media 目录
4. **数据库索引**: 在 created_at 和 user 字段上建立索引
5. **查询优化**: 使用 select_related 减少数据库查询

## 扩展性

### 添加新功能

1. **新的模板类型**
   - 在 `template/` 添加新模板
   - 更新 `forms.py` 中的选项
   - 更新 `template.json` 模板定义

2. **新的配置选项**
   - 在 `models.py` 添加字段
   - 在 `forms.py` 添加表单字段
   - 在模板中添加 UI
   - 创建数据库迁移

3. **新的处理逻辑**
   - 在 `views.py` 添加视图函数
   - 在 `urls.py` 添加路由
   - 创建对应的模板
   - 添加权限检查（如需要）

4. **新的用户角色**
   - 在 `init_users.py` 添加组和权限
   - 在 `models.py` 添加自定义权限
   - 在视图中添加权限装饰器
   - 在模板中添加条件判断

### 部署建议

1. **生产环境**
   - 使用 PostgreSQL 或 MySQL
   - 使用 Gunicorn + Nginx
   - 启用 HTTPS
   - 配置静态文件服务
   - 设置环境变量（SECRET_KEY、ALLOWED_HOSTS）
   - 配置全局 LLM 设置

2. **性能优化**
   - 使用 Redis 缓存
   - 使用 Celery 异步任务队列
   - 配置 CDN 加速静态文件
   - 启用数据库连接池

3. **监控**
   - 配置日志记录
   - 使用 Sentry 错误追踪
   - 监控服务器资源
   - 监控 LLM API 调用次数和成本

4. **备份**
   - 定期备份数据库
   - 备份用户上传的文件
   - 备份全局配置

## 目录结构说明

```
web/
├── manage.py                      # Django 管理命令
├── db.sqlite3                    # SQLite 数据库
├── web_frontend/                 # 项目配置
│   ├── settings.py               # 配置文件
│   ├── urls.py                   # 主路由
│   ├── wsgi.py                   # WSGI 入口
│   └── context_processors.py    # 自定义上下文处理器
├── ppt_generator/                # 应用
│   ├── models.py                 # 数据模型
│   │   ├── PPTGeneration         # PPT 生成记录
│   │   └── GlobalLLMConfig       # 全局 LLM 配置（单例）
│   ├── views.py                  # 视图逻辑
│   ├── forms.py                  # 表单定义
│   ├── urls.py                   # 应用路由
│   ├── admin.py                  # 管理后台
│   ├── migrations/               # 数据库迁移
│   └── management/               # 管理命令
│       └── commands/
│           └── init_users.py     # 初始化用户和权限
├── templates/                    # HTML 模板
│   ├── base.html                 # 基础模板
│   ├── login.html                # 登录页面
│   └── ppt_generator/            # 应用模板
│       ├── index.html            # 首页（上传表单）
│       ├── detail.html           # 生成详情页
│       ├── history.html          # 历史记录
│       └── developer_tools.html # 开发者工具（开发者）
├── static/                       # 静态文件
│   ├── css/
│   │   └── style.css             # 主样式表
│   ├── js/
│   │   └── main.js               # 主脚本
│   └── images/                   # 图片资源
└── media/                        # 用户文件
    ├── scripts/                  # 上传的 DOCX
    ├── templates/                # 上传的 PPTX 模板
    ├── outputs/                  # 生成的 PPT
    └── configs/                  # 生成的 JSON
```

## 数据库模型关系

```
User (Django 内置)
  ├─ 1:N → PPTGeneration (用户的生成记录)
  └─ 1:1 → GlobalLLMConfig.updated_by (配置更新者)

Group (Django 内置)
  └─ 开发者组 (拥有特殊权限)

GlobalLLMConfig (单例)
  └─ 全局 LLM 配置，所有用户共享

PPTGeneration
  ├─ user (ForeignKey → User)
  ├─ docx_file (FileField)
  ├─ template_file (FileField)
  ├─ output_ppt (FileField)
  ├─ output_json (FileField)
  └─ LLM 配置字段（可覆盖全局配置）
```

## 关键文件说明

### models.py

- **PPTGeneration**: 存储每次生成任务的信息
  - 用户关联、文件路径、状态、LLM 配置等
  - 提供 `mark_processing()`, `mark_completed()`, `mark_failed()` 方法

- **GlobalLLMConfig**: 全局 LLM 配置（单例模式）
  - 提供 `get_config()` 类方法获取配置
  - 限制只能有一个实例

### views.py

- **index()**: 首页，显示上传表单和最近记录
- **generation_detail()**: 生成详情页，显示状态和下载链接
- **start_generation()**: AJAX 端点，启动生成任务
  - 获取全局 LLM 配置作为 fallback
  - 调用后端脚本生成 PPT
- **check_status()**: AJAX 端点，检查生成状态
- **history()**: 历史记录页面（按用户过滤）
- **download_ppt/json()**: 文件下载
- **developer_tools()**: 开发者工具主页面（开发者权限）
  - 双 Tab 界面：生成配置模板 / 编辑配置模板
- **generate_config_template()**: AJAX 端点，从 PPTX 生成配置模板
- **ai_enrich_template_view()**: AJAX 端点，使用 AI 填充配置模板

### admin.py

- **GlobalLLMConfigAdmin**: 全局配置管理界面
  - 单例模式，不允许添加/删除
  - 记录更新者和更新时间

- **PPTGenerationAdmin**: 生成记录管理界面
  - 自定义列表显示、过滤器、搜索
  - 彩色状态徽章、下载链接

### context_processors.py

- **user_role_processor()**: 提供 `is_developer` 变量给所有模板
