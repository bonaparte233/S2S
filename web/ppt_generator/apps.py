"""
App configuration for PPT Generator.
"""
from django.apps import AppConfig


class PptGeneratorConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'ppt_generator'
    verbose_name = 'PPT生成器'

