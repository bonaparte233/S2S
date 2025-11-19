"""
Management command to initialize default users and groups.
"""

from django.core.management.base import BaseCommand
from django.contrib.auth.models import User, Group, Permission
from django.contrib.contenttypes.models import ContentType
from ppt_generator.models import GlobalLLMConfig, PPTGeneration


class Command(BaseCommand):
    help = "初始化默认用户和权限组"

    def handle(self, *args, **options):
        self.stdout.write("🔧 初始化用户和权限...")

        # 创建开发者组
        developer_group, created = Group.objects.get_or_create(name="开发者")
        if created:
            self.stdout.write(self.style.SUCCESS("✅ 创建开发者组"))
        else:
            self.stdout.write("ℹ️  开发者组已存在")

        # 总是更新开发者权限（无论组是否新创建）
        content_type = ContentType.objects.get_for_model(PPTGeneration)
        permissions = Permission.objects.filter(
            content_type=content_type,
            codename__in=[
                "is_developer",
                "can_export_template_json",
                "can_view_llm_config",
            ],
        )
        developer_group.permissions.set(permissions)
        self.stdout.write(
            self.style.SUCCESS(f"✅ 配置开发者权限（共{permissions.count()}个）")
        )

        # 创建管理员账户
        admin_username = "admin"
        admin_password = "admin123"
        if not User.objects.filter(username=admin_username).exists():
            admin = User.objects.create_superuser(
                username=admin_username,
                email="admin@s2s.local",
                password=admin_password,
            )
            admin.groups.add(developer_group)
            self.stdout.write(
                self.style.SUCCESS(
                    f"✅ 创建管理员账户: {admin_username} / {admin_password}"
                )
            )
        else:
            self.stdout.write(f"ℹ️  管理员账户已存在: {admin_username}")

        # 创建默认普通用户
        user_username = "user"
        user_password = "user123"
        if not User.objects.filter(username=user_username).exists():
            user = User.objects.create_user(
                username=user_username, email="user@s2s.local", password=user_password
            )
            self.stdout.write(
                self.style.SUCCESS(
                    f"✅ 创建普通用户账户: {user_username} / {user_password}"
                )
            )
        else:
            self.stdout.write(f"ℹ️  普通用户账户已存在: {user_username}")

        # 创建默认开发者用户
        dev_username = "developer"
        dev_password = "dev123"
        if not User.objects.filter(username=dev_username).exists():
            developer = User.objects.create_user(
                username=dev_username,
                email="developer@s2s.local",
                password=dev_password,
            )
            developer.groups.add(developer_group)
            self.stdout.write(
                self.style.SUCCESS(
                    f"✅ 创建开发者账户: {dev_username} / {dev_password}"
                )
            )
        else:
            self.stdout.write(f"ℹ️  开发者账户已存在: {dev_username}")

        # 创建全局LLM配置
        global_config = GlobalLLMConfig.get_config()
        self.stdout.write(
            self.style.SUCCESS(
                f"✅ 全局LLM配置已就绪: {global_config.llm_provider} - {global_config.llm_model}"
            )
        )

        self.stdout.write(self.style.SUCCESS("\n🎉 用户初始化完成！"))
        self.stdout.write("\n默认账户：")
        self.stdout.write(f"  管理员: {admin_username} / {admin_password}")
        self.stdout.write(f"  开发者: {dev_username} / {dev_password}")
        self.stdout.write(f"  普通用户: {user_username} / {user_password}")
        self.stdout.write("\n💡 提示：请在Admin后台配置全局LLM的API密钥")
