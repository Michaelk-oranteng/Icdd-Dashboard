#!/usr/bin/env bash
# Exit immediately if any command fails
set -o errexit

echo "==> Installing Python dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

echo "==> Collecting static files..."
python manage.py collectstatic --no-input

echo "==> Running database migrations..."
python manage.py migrate

echo "==> Build complete."