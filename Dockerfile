FROM python:3.10-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# WeasyPrint runtime dependencies (replaces LibreOffice entirely)
# pango + cairo = text layout + rendering
# fonts-liberation2 = Arial/Times New Roman substitutes for document fonts
# fonts-indic / fonts-noto = coverage for Indian language characters
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangoft2-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    shared-mime-info \
    fonts-liberation2 \
    fonts-dejavu-core \
    fonts-indic \
    fonts-noto-core \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Custom agency fonts (Times New Roman etc. from your fonts/ folder)
RUN mkdir -p /usr/share/fonts/truetype/custom
COPY ./fonts/ /usr/share/fonts/truetype/custom/
RUN fc-cache -f -v

WORKDIR /app
COPY requirements.txt /app/
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . /app/

EXPOSE 5000

# WeasyPrint is in-process — no LibreOffice RAM spikes
# Multiple workers are now safe: each WeasyPrint render uses ~50 MB peak
# Timeout 60s is plenty — 5 PDFs via WeasyPrint take ~5-10s total
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "2", \
     "--threads", "4", \
     "--timeout", "60", \
     "--keep-alive", "5", \
     "app:app"]