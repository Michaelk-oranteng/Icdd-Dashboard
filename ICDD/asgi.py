"""
ASGI config for ICDD project.
"""

import os

from django.core.asgi import get_asgi_application

SETTINGS_MODULE = (
    'ICDD.deployment_settings'
    if "RENDER_EXTERNAL_HOSTNAME" in os.environ
    else 'ICDD.settings'
)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', SETTINGS_MODULE)

application = get_asgi_application()