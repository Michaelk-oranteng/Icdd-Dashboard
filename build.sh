#!/usr/bin/env bash
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt

python manage.py collectstatic --no-input
python manage.py migrate

# Optional: auto-create superuser from env vars
# python manage.py createsuperuser --no-input || true