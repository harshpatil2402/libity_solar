# ── Stage: production image ────────────────────────────────────────────────────
FROM python:3.10-slim

# Keeps Python from writing .pyc files and forces stdout/stderr to be unbuffered
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    # Fixed path for Playwright browser binaries so they survive image layers
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

# ── 1. System dependencies ─────────────────────────────────────────────────────
# WeasyPrint: Pango/Cairo rendering stack + shared-mime-info
# Fonts: Liberation (latin), Indic (Devanagari/Marathi), Noto core (Unicode fallback)
# gcc is required for some Python package C extensions during pip install
# NOTE: apt lists are intentionally NOT cleaned here so that the next stage's
#       `playwright install-deps` can call apt-get without a separate update step.
RUN apt-get update && apt-get install -y --no-install-recommends \
    # WeasyPrint rendering
    libpango-1.0-0 \
    libpangoft2-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    shared-mime-info \
    # Fonts
    fonts-liberation \
    fonts-indic \
    fonts-noto-core \
    # Build tools
    gcc

# ── 2. Python dependencies ─────────────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── 3. Playwright Chromium ─────────────────────────────────────────────────────
# `install chromium` fetches the browser binary to PLAYWRIGHT_BROWSERS_PATH.
# `install-deps chromium` installs its OS-level shared libraries via apt-get;
#   this works because the apt lists from step 1 are still accessible in the
#   lower overlay layer. Lists are purged here to minimise the final image size.
RUN playwright install chromium \
    && playwright install-deps chromium \
    && rm -rf /var/lib/apt/lists/*

# ── 4. Application code ────────────────────────────────────────────────────────
COPY . .

# ── 5. Runtime ─────────────────────────────────────────────────────────────────
EXPOSE 5000

# CRITICAL — single worker:
#   The PDF job queue uses an in-memory `jobs` dict in utils/doc_engine.py.
#   Multiple workers each have their own memory space; a request that starts
#   a job on worker A cannot be polled from worker B.  One worker + multiple
#   threads is the correct model here.
#
# Runtime env vars must be injected at `docker run` time (not baked in):
#   docker run --env-file .env ...
#   or individual -e flags for SECRET_KEY, SUPABASE_URL, SUPABASE_KEY, etc.
#
# WhatsApp Web session (wa_profile/) is ephemeral inside the container.
#   Mount a volume to persist it:
#   docker run -v /host/wa_profile:/app/wa_profile ...
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "1", \
     "--threads", "8", \
     "--worker-class", "gthread", \
     "--timeout", "120", \
     "run:app"]
