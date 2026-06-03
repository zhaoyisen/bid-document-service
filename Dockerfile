FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    APP_ENV=production \
    HOST=0.0.0.0 \
    PORT=8010 \
    OUTPUT_ROOT=/app/outputs \
    MAX_UPLOAD_MB=100 \
    OUTPUT_RETENTION_HOURS=168 \
    PUBLIC_DOWNLOAD_TTL_SECONDS=86400

WORKDIR /app

RUN useradd --create-home --shell /usr/sbin/nologin appuser

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY bid_document_service /app/bid_document_service
COPY examples /app/examples
COPY README.md /app/README.md

RUN mkdir -p /app/outputs && chown -R appuser:appuser /app

USER appuser

EXPOSE 8010

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8010/health', timeout=3).read()"

CMD ["python", "-m", "uvicorn", "bid_document_service.main:app", "--host", "0.0.0.0", "--port", "8010", "--workers", "2"]
