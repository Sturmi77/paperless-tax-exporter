FROM python:3.12-slim

LABEL maintainer="paperless-tax-exporter"
LABEL description="Steuerberater-Export für Paperless NGX"

WORKDIR /app

# Abhängigkeiten zuerst (Layer-Cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App-Dateien (ARG vor COPY erzwingt Cache-Invalidierung bei jedem Build)
# COPY . . kopiert alle Dateien – keine manuelle Liste noetig bei neuen Modulen
ARG CACHEBUST=1
COPY . .

# Ausgabe-Verzeichnis anlegen
RUN mkdir -p /output

EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "3600", "app:app"]
