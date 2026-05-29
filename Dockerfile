FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=3000 \
    GENOMA_DATA_DIR=/data \
    CHROME_PATH=/usr/bin/chromium

RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fonts-liberation \
    fonts-dejavu \
    libnss3 \
    libxss1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY genoma_app ./genoma_app
COPY GENOMA_NINOS3_extracted ./GENOMA_NINOS3_extracted

RUN mkdir -p /data/uploads

EXPOSE 3000

CMD ["python", "-u", "genoma_app/app.py"]
