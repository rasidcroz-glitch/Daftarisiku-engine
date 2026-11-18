# DaftarIsiKu engine — Python + LibreOffice headless.
# Base image kecil (slim), LibreOffice Writer saja (bukan paket penuh)
# + font Liberation agar metrik Times New Roman akurat saat render PDF.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-liberation \
        fonts-dejavu \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY toc.py server.py gunicorn.conf.py ./
# results/ dibuat saat runtime; jangan bawa hasil job lokal ke image.
RUN mkdir -p results && useradd -m appuser && chown -R appuser:appuser /app
USER appuser

# Render & Fly.io memberi PORT lewat env; default 5000.
ENV PORT=5000
EXPOSE 5000

CMD ["gunicorn", "-c", "gunicorn.conf.py", "server:app"]
