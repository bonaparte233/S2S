@echo off
chcp 65001
set PYTHONIOENCODING=utf-8
cd web
python manage.py migrate
python manage.py runserver 0.0.0.0:8000

