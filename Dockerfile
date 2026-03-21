FROM python:3.12-slim

LABEL maintainer="paperless-tax-exporter"
LABEL description="Steuerberater-Export für Paperless NGX"

WORKDIR /app

# Abhängigkeiten zuerst (Layer-Cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App-Dateien
COPY app.py excel_export.py pdf_export.py ./
COPY templates/ templates/
COPY static/ static/

# Ausgabe-Verzeichnis anlegen
RUN mkdir -p /output

EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "300", "app:app"]
