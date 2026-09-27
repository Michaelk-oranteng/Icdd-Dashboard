"""
Deployment settings — production overrides on Render.

Loaded only when RENDER_EXTERNAL_HOSTNAME is present in the environment
(see manage.py / wsgi.py).
"""

import os

import dj_database_url

from .settings import *          # noqa: F401,F403
from .settings import BASE_DIR   # noqa: F401


# ============================================================
# CORE
# ============================================================

DEBUG = False

SECRET_KEY = os.environ.get('SECRET_KEY')


# ============================================================
# HOSTS / CSRF
# ============================================================
# RENDER_EXTERNAL_HOSTNAME is auto-injected by Render, e.g.
#   omnibsic-dashboard.onrender.com

RENDER_HOSTNAME = os.environ.get('RENDER_EXTERNAL_HOSTNAME', '').strip()

ALLOWED_HOSTS = [
    h for h in [
        RENDER_HOSTNAME,
        '.onrender.com',        # covers all Render subdomains
        'localhost',
        '127.0.0.1',
    ] if h
]

CSRF_TRUSTED_ORIGINS = [
    o for o in [
        f'https://{RENDER_HOSTNAME}' if RENDER_HOSTNAME else '',
        'https://*.onrender.com',
        'http://localhost:8000',
        'http://127.0.0.1:8000',
    ] if o
]


# ============================================================
# MIDDLEWARE
# ============================================================
# (Identical to settings.py — restated here for clarity / future edits)

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]


# ============================================================
# STORAGE
# ============================================================
# NOTE: the correct key is STORAGES (plural) in Django 5+.

STORAGES = {
    'default': {
        'BACKEND': 'django.core.files.storage.FileSystemStorage',
    },
    'staticfiles': {
        'BACKEND': 'whitenoise.storage.CompressedStaticFilesStorage',
    },
}


# ============================================================
# DATABASE
# ============================================================
# Render injects DATABASE_URL automatically when you link a
# PostgreSQL instance to the web service.

DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()

if DATABASE_URL:
    DATABASES = {
        'default': dj_database_url.parse(
            DATABASE_URL,
            conn_max_age=600,
            conn_health_checks=True,
        )
    }
else:
    # Fallback — no DATABASE_URL (shouldn't happen on Render once you
    # attach Postgres, but keeps local `runserver` working).
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.sqlite3',
            'NAME': BASE_DIR / 'db.sqlite3',
        }
    }