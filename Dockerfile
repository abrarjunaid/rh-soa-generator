FROM python:3.12-slim-bookworm

# Install ALL Chromium dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libnss3 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    libxshmfence1 \
    libx11-xcb1 \
    libxcb-dri3-0 \
    libdbus-1-3 \
    libatspi2.0-0 \
    libxfixes3 \
    libxext6 \
    libx11-6 \
    libxcb1 \
    libxau6 \
    libxdmcp6 \
    libxrender1 \
    libxi6 \
    libxtst6 \
    libglib2.0-0 \
    libnspr4 \
    libexpat1 \
    fonts-liberation \
    fonts-unifont \
    xdg-utils \
    wget \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium only (deps already installed above)
RUN playwright install chromium

COPY . .

ENV PORT=5000
EXPOSE $PORT

CMD gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 1 app:app
