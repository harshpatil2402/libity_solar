FROM python:3.10-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# Enable contrib repo (needed for ttf-mscorefonts-installer)
RUN if [ -f /etc/apt/sources.list.d/debian.sources ]; then \
        sed -i 's/Components: main/Components: main contrib/g' /etc/apt/sources.list.d/debian.sources; \
    fi && \
    if [ -f /etc/apt/sources.list ]; then \
        sed -i 's/\bmain\b/main contrib/g' /etc/apt/sources.list; \
    fi

RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    ttf-mscorefonts-installer \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Custom fonts
RUN mkdir -p /usr/share/fonts/truetype/custom
COPY ./fonts/ /usr/share/fonts/truetype/custom/
RUN fc-cache -f -v

WORKDIR /app
COPY requirements.txt /app/
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . /app/

EXPOSE 5000

# Single worker process — prevents double LibreOffice RAM usage
# 4 threads — handles concurrent HTTP requests within the single worker
# 180s timeout — LibreOffice PDF conversion for 5 docs can take ~60-90s
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "1", \
     "--threads", "4", \
     "--timeout", "180", \
     "--keep-alive", "5", \
     "app:app"]