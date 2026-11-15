# One container runs both agent.py (Flask) and bridge.js (WhatsApp via headless Chromium).
FROM node:20-bookworm-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        chromium fonts-liberation fonts-noto-color-emoji \
        python3 python3-venv tini ca-certificates \
    && rm -rf /var/lib/apt/lists/*

ENV PUPPETEER_SKIP_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium \
    PYTHONUNBUFFERED=1 \
    LAB_EXCEL=/data/lab_inventory.xlsx \
    WWEBJS_AUTH_DIR=/data/wwebjs_auth \
    PORT=8080

WORKDIR /app

COPY package.json ./
RUN npm install --omit=dev && npm cache clean --force

COPY requirements.txt ./
RUN python3 -m venv /opt/venv && /opt/venv/bin/pip install --no-cache-dir -r requirements.txt
ENV PATH="/opt/venv/bin:$PATH"

COPY agent.py bridge.js ./

# /data must be a persistent volume: it holds the Excel database and the WhatsApp login.
VOLUME ["/data"]
EXPOSE 8080

ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["python", "agent.py"]
