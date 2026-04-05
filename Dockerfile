FROM python:3.12-slim

# Install system dependencies for Playwright/Chromium
RUN apt-get update && apt-get install -y --no-install-recommends \
    libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 \
    libxkbcommon0 libxcomposite1 libxdamage1 libxrandr2 libgbm1 \
    libpango-1.0-0 libcairo2 libasound2 libxshmfence1 \
    libx11-xcb1 libxfixes3 libxext6 libx11-6 libxcb1 \
    fonts-liberation wget ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium
RUN playwright install chromium --with-deps

# Copy app
COPY . .

# Railway sets PORT env var automatically
ENV PORT=5000

# Run with gunicorn - bind to 0.0.0.0 on Railway's PORT
CMD gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 2 app:app
