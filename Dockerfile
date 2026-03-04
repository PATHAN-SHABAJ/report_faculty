# Use python 3.11 slim image
FROM python:3.11-slim

# Install LibreOffice for PDF conversion and clean up apt cache
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set the working directory
WORKDIR /app

# Copy requirements file first for caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Expose Render's standard port
EXPOSE 10000

# Run gunicorn using the PORT environment variable Render provides
CMD gunicorn app:app --bind 0.0.0.0:${PORT:-10000}
