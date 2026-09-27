# ICDD/urls.py

import os

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from django.http import Http404
from django.views.static import serve as serve_static


def serve_media(request, path):
    """Serve media files (avatars) in both DEBUG and production."""
    full_path = os.path.join(str(settings.MEDIA_ROOT), path)
    if not os.path.exists(full_path):
        raise Http404('Media file not found.')
    return serve_static(request, path, document_root=settings.MEDIA_ROOT)


urlpatterns = [
    path('admin/', admin.site.urls),
    path('adminboard/', include('control_dashboard.urls')),

    # Media files served at the root — matches MEDIA_URL = '/media/'
    path('media/<path:path>', serve_media, name='serve_media'),
]