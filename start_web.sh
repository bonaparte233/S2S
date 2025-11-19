#!/bin/bash

# S2S Web 应用启动脚本

echo "🚀 启动 S2S Web 应用..."

# 检查虚拟环境
if [ ! -d ".venv" ]; then
    echo "⚠️  未找到虚拟环境，正在创建..."
    python3 -m venv .venv
    source .venv/bin/activate
    echo "📥 安装依赖..."
    pip install -r requirements.txt
else
    source .venv/bin/activate
fi

# 进入 web 目录
cd web

# 运行数据库迁移
echo "🗄️  运行数据库迁移..."
python manage.py makemigrations
python manage.py migrate

# 初始化默认用户
echo "👥 初始化默认用户..."
python manage.py init_users

# 启动开发服务器
echo ""
echo "✅ 启动开发服务器..."
echo "🌐 访问地址: http://127.0.0.1:8000/"
echo "🛑 按 Ctrl+C 停止服务器"
echo ""
python manage.py runserver

