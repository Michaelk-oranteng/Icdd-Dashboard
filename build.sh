#!/usr/bin/env bash
# Exit on error
set -o errexit

# Install dependencies
pip install --upgrade pip
pip install -r requirements.txt

# Collect static files into STATIC_ROOT
python manage.py collectstatic --no-input

# Apply database migrations
python manage.py migrate

# Optional: create a superuser automatically from env vars
# (uncomment if you set DJANGO_SUPERUSER_* env vars on Render)
# python manage.py createsuperuser --no-input || true