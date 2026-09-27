"""
WSGI config for ICDD project.

It exposes the WSGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/6.0/howto/deployment/wsgi/
"""

import os

from django.core.wsgi import get_wsgi_application

SETTINGS_MODULE = 'api.deployment_settings' if "RENDER_EXTERNAL_HOSTNAME" in os.environ else 'api.settings'
os.environ.setdefault('DJANGO_SETTINGS_MODULE', SETTINGS_MODULE)    

application = get_wsgi_application()
