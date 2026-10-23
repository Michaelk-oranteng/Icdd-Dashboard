"""
Django settings for ICDD project.

Configuration is entirely environment-driven so the same file works
locally, on Railway, and anywhere else without modification.
"""

import os
from pathlib import Path

import dj_database_url

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR = Path(__file__).resolve().parent.parent


# ============================================================
# CORE
# ============================================================

DEBUG = os.environ.get('DEBUG', 'False').lower() in ('1', 'true', 'yes')

SECRET_KEY = os.environ.get('SECRET_KEY')

if not SECRET_KEY:
    if DEBUG:
        # Local dev only — never used in production.
        SECRET_KEY = 'django-insecure-local-dev-only-do-not-use-in-production'
    else:
        raise RuntimeError(
            "SECRET_KEY environment variable is required in production. "
            "Generate one with: python -c "
            "\"from django.core.management.utils import get_random_secret_key; "
            "print(get_random_secret_key())\""
        )


# ============================================================
# HOSTS / CSRF
# ============================================================
# Override these in production by setting the ALLOWED_HOSTS and
# CSRF_TRUSTED_ORIGINS environment variables.
#
#   ALLOWED_HOSTS=.up.railway.app,mycustomdomain.com
#   CSRF_TRUSTED_ORIGINS=https://*.up.railway.app,https://mycustomdomain.com

ALLOWED_HOSTS = [
    h.strip() for h in os.environ.get(
        'ALLOWED_HOSTS',
        'localhost,127.0.0.1,.up.railway.app,.railway.app'
    ).split(',') if h.strip()
]

CSRF_TRUSTED_ORIGINS = [
    o.strip() for o in os.environ.get(
        'CSRF_TRUSTED_ORIGINS',
        'http://localhost:8000,http://127.0.0.1:8000,'
        'https://*.up.railway.app,https://*.railway.app'
    ).split(',') if o.strip()
]


# ============================================================
# AUTH REDIRECTS
# ============================================================

LOGIN_URL = 'control_dashboard:login'
LOGIN_REDIRECT_URL = 'control_dashboard:landing_page'


# ============================================================
# APPLICATIONS
# ============================================================

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'control_dashboard',
    'import_export',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',

    # ═══ 10-MINUTE INACTIVITY LOGOUT ═══
    # Must be placed LAST so it runs after AuthenticationMiddleware
    # has already populated request.user.
    'control_dashboard.middleware.IdleTimeoutMiddleware',
]

ROOT_URLCONF = 'ICDD.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'ICDD.wsgi.application'


# ============================================================
# DATABASE
# ============================================================
# If DATABASE_URL is set → connect to that (PostgreSQL in production).
# Otherwise           → fall back to local SQLite.

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
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.sqlite3',
            'NAME': BASE_DIR / 'db.sqlite3',
        }
    }


# ============================================================
# PASSWORD VALIDATION
# ============================================================

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]


# ============================================================
# I18N
# ============================================================

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True


# ============================================================
# STATIC FILES
# ============================================================

STATIC_URL = '/static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

STORAGES = {
    'default': {
        'BACKEND': 'django.core.files.storage.FileSystemStorage',
    },
    'staticfiles': {
        'BACKEND': 'whitenoise.storage.CompressedStaticFilesStorage',
    },
}


# ============================================================
# UPLOAD LIMITS
# ============================================================

DATA_UPLOAD_MAX_MEMORY_SIZE = 50 * 1024 * 1024   # 50 MB
FILE_UPLOAD_MAX_MEMORY_SIZE = 50 * 1024 * 1024   # 50 MB


# ============================================================
# SESSION / INACTIVITY
# ============================================================
# ═══ 10-MINUTE INACTIVITY POLICY ═══
#
# The three layers below cooperate to log the user out after
# 10 minutes of idle time — regardless of tab state, JS tampering,
# or direct cookie replay:
#
#   1. SESSION_COOKIE_AGE           → 10-minute cookie lifetime
#   2. SESSION_SAVE_EVERY_REQUEST   → refresh the cookie on each hit
#                                     (so "active" users stay logged in)
#   3. IdleTimeoutMiddleware        → server-side check of a
#                                     `_last_activity` timestamp
#   4. base.html JS timer           → client-side countdown + banner
#
# All four must agree: 600 seconds = 10 minutes.

INACTIVITY_TIMEOUT_SECONDS = int(os.environ.get('INACTIVITY_TIMEOUT_SECONDS', '600'))  # 10 min

SESSION_COOKIE_AGE = INACTIVITY_TIMEOUT_SECONDS
SESSION_SAVE_EVERY_REQUEST = True
SESSION_EXPIRE_AT_BROWSER_CLOSE = True
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = 'Lax'

# In local dev over HTTP, cookies cannot be Secure.
# In production over HTTPS, both of these become True automatically.
SESSION_COOKIE_SECURE = not DEBUG
CSRF_COOKIE_SECURE = not DEBUG
CSRF_COOKIE_HTTPONLY = False
CSRF_COOKIE_SAMESITE = 'Lax'


# ============================================================
# SECURITY
# ============================================================

SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')

# Set DISABLE_SSL_REDIRECT=True in the environment only if the
# platform's health checks get caught in a redirect loop.
SECURE_SSL_REDIRECT = (
    not DEBUG
    and os.environ.get('DISABLE_SSL_REDIRECT', 'False').lower() not in ('1', 'true', 'yes')
)

SECURE_HSTS_SECONDS = 0 if DEBUG else 60 * 60 * 24 * 30
SECURE_HSTS_INCLUDE_SUBDOMAINS = not DEBUG
SECURE_HSTS_PRELOAD = False
SECURE_CONTENT_TYPE_NOSNIFF = True
SECURE_REFERRER_POLICY = 'same-origin'
X_FRAME_OPTIONS = 'DENY'


# ============================================================
# LOGGING
# ============================================================

LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose': {
            'format': '{levelname} {asctime} {module} {message}',
            'style': '{',
        },
    },
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'verbose',
        },
    },
    'root': {
        'handlers': ['console'],
        'level': os.environ.get('LOG_LEVEL', 'INFO'),
    },
    'loggers': {
        'django.request': {
            'handlers': ['console'],
            'level': 'WARNING',
            'propagate': False,
        },
        'control_dashboard': {
            'handlers': ['console'],
            'level': os.environ.get('LOG_LEVEL', 'INFO'),
            'propagate': False,
        },
    },
}


# ============================================================
# DEFAULT PRIMARY KEY
# ============================================================

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'