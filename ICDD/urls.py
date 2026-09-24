# ICDD/urls.py

from django.contrib import admin
from django.urls import path, include
from django.shortcuts import redirect

urlpatterns = [
    path('', lambda request: redirect('control_dashboard:landing_page'), name='home'),
    path('admin/', admin.site.urls),
    path('adminboard/', include('control_dashboard.urls')),  # This should be correct
]