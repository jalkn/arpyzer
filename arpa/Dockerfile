# --- STAGE 1: BUILDER & PRODUCTION DEPENDENCIES ---
FROM python:3.11-slim as builder

# Install system dependencies needed for libraries like PyMuPDF and for building Python packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libopenjp2-7 \
    libtiff5 \
    libpng-dev \
    zlib1g-dev \
    # Clean up APT lists to keep the image small
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONUNBUFFERED 1
ENV DJANGO_SETTINGS_MODULE arpa.settings
# Cloud Run typically expects port 8080
ENV PORT 8080

# Create and set working directory
WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project code
COPY . .

# Collect static files (needed for Whitenoise to serve them)
RUN python3 manage.py collectstatic --no-input

# --- STAGE 2: FINAL MINIMAL IMAGE ---
# Use a highly minimal runtime image
FROM python:3.11-slim

WORKDIR /app

# Copy runtime files from the builder stage
COPY --from=builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=builder /app /app

# Cloud Run exposes port 8080
EXPOSE 8080

# Start the application using Gunicorn, binding to 0.0.0.0 and the PORT environment variable
CMD ["gunicorn", "arpa.wsgi:application", "--bind", "0.0.0.0:8080", "--workers", "2", "--timeout", "120"]
