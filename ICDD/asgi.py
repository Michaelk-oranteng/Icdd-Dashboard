"""
ASGI config for ICDD project.

It exposes the ASGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/6.0/howto/deployment/asgi/
"""

import os

from django.core.asgi import get_asgi_application

SETTINGS_MODULE = 'api.deployment_settings' if "RENDER_EXTERNAL_HOSTNAME" in os.environ else 'api.settings'
os.environ.setdefault('DJANGO_SETTINGS_MODULE', SETTINGS_MODULE)

application = get_asgi_application()
