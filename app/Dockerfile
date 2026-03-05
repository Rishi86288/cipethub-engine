FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    ffmpeg \
    libreoffice-impress \
    fonts-dejavu-core \
    fonts-noto \
    fonts-noto-color-emoji \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p /videos

EXPOSE 8899
HEALTHCHECK --interval=30s --timeout=10s CMD curl -f http://localhost:8899/health || exit 1

CMD ["python3", "server.py"]
