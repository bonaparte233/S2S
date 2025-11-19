"""
WSGI config for S2S web project.
"""

import os

from django.core.wsgi import get_wsgi_application

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 's2s_web.settings')

application = get_wsgi_application()

