FROM python:3.11-slim

# ── System deps needed by Playwright/Chromium ────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget gnupg ca-certificates \
        fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ── Python deps ───────────────────────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Playwright: install Chromium + its system libraries ───────────────────────
RUN playwright install chromium --with-deps

# ── App source ────────────────────────────────────────────────────────────────
COPY main.py converter.py mermaid_renderer.py ./
COPY static/ static/

# ── Runtime ──────────────────────────────────────────────────────────────────
EXPOSE 8765
ENV PYTHONUNBUFFERED=1

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8765"]
