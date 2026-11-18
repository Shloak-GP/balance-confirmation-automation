# Balance-check tool as one isolated container (for the GP Assist tool marketplace / a plain server).
FROM python:3.11-slim
RUN useradd -m -u 10001 app
WORKDIR /app
COPY requirements.txt .
# pywin32 is Windows-only and is skipped on Linux by its marker
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV BC_CONFIG=/data/config.json PYTHONUNBUFFERED=1
RUN mkdir -p /data && chown -R app /data /app
USER app
EXPOSE 8080
HEALTHCHECK --interval=60s --timeout=5s CMD python -c "import urllib.request,json;d=json.load(urllib.request.urlopen('http://127.0.0.1:8080/health'));raise SystemExit(0 if d['all_ok'] else 1)"
CMD ["python", "web.py"]
