#!/bin/sh

# Run database migrations
python manage.py migrate

# Start Gunicorn
gunicorn --bind 0.0.0.0:8080 --workers 2 arpa.wsgi:application
