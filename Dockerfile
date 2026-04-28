FROM python:3.11-slim

LABEL maintainer="Arthur Motta"
LABEL description="Crítica da Base Cadastral EFPC — Plotly Dash"

RUN apt-get update && apt-get install -y --no-install-recommends gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/  ./src/
COPY app/  ./app/

RUN mkdir -p data/raw data/processed results/reports

RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

EXPOSE 8050

HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8050')" \
    || exit 1

CMD ["gunicorn", "app.dashboard:server", \
     "--bind", "0.0.0.0:8050", \
     "--workers", "2", \
     "--timeout", "120", \
     "--access-logfile", "-"]
