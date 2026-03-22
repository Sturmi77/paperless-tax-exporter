import os
import re
import io
import json
import requests
import threading
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, send_file

from excel_export import create_excel
from pdf_export import download_pdfs

app = Flask(__name__)

# Cache-Busting: Static-Files bekommen einen Versions-Query-Parameter
# basierend auf dem Build-Zeitpunkt – verhindert Browser-Cache-Probleme
import time as _time
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0  # kein Browser-Caching
_BUILD_VERSION = str(int(_time.time()))

@app.context_processor
def inject_version():
    return {"ver": _BUILD_VERSION}

PAPERLESS_URL   = os.environ.get("PAPERLESS_URL",   "http://192.168.178.115:8000")
PAPERLESS_TOKEN = os.environ.get("PAPERLESS_TOKEN", "")
OUTPUT_DIR      = os.environ.get("OUTPUT_DIR",      "/output")

# ── Sicherheit: nur GET-Zugriffe auf Paperless (read-only) ─────────────
# Der Container hat keinen Schreibzugriff auf Paperless.
# Schreibzugriff existiert ausschließlich auf OUTPUT_DIR.

ALLOWED_PAPERLESS_METHODS = {"GET"}  # Wird in paperless_get() erzwungen

def _assert_output_path(path: str):
    """Stellt sicher, dass Schreibzugriffe nur innerhalb OUTPUT_DIR erfolgen."""
    real_output = os.path.realpath(OUTPUT_DIR)
    real_path   = os.path.realpath(path)
    if not real_path.startswith(real_output):
        raise PermissionError(
            f"Schreibzugriff außerhalb des erlaubten Pfads verweigert: {path}"
        )

# Globaler Job-Status
job_status = {
    "running": False,
    "log": [],
    "done": False,
    "error": None,
    "excel_path": None,
    "doc_count": 0,
}
job_lock = threading.Lock()


def paperless_get(path, params=None):
    """Einziger HTTP-Zugang zu Paperless – ausschließlich GET (read-only)."""
    headers = {"Authorization": f"Token {PAPERLESS_TOKEN}"}
    # Sicherstellen dass der Pfad auf /api/ zeigt
    if not path.startswith("/api/"):
        raise ValueError(f"Ungültiger API-Pfad: {path}")
    url = f"{PAPERLESS_URL.rstrip('/')}{path}"
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_all_tags():
    tags = []
    url = "/api/tags/"
    while url:
        data = paperless_get(url)
        tags.extend(data.get("results", []))
        next_url = data.get("next")
        if next_url:
            # Nur den Pfad weitergeben
            url = next_url.replace(PAPERLESS_URL.rstrip("/"), "")
        else:
            url = None
    return tags


def get_documents(date_from, date_to, tag_ids):
    """Holt alle Dokumente gefiltert nach Datum und Tags (paginiert)."""
    documents = []
    params = {
        "created__date__gte": date_from,
        "created__date__lte": date_to,
        "page_size": 100,
    }
    if tag_ids:
        params["tags__id__all"] = ",".join(str(t) for t in tag_ids)

    path = "/api/documents/"
    page = 1
    while path:
        params["page"] = page
        data = paperless_get(path, params=params)
        results = data.get("results", [])
        documents.extend(results)
        if data.get("next"):
            page += 1
        else:
            break
    return documents


def run_export_job(date_from, date_to, tag_ids, tag_names, year_label):
    global job_status
    try:
        with job_lock:
            job_status["log"] = []
            job_status["done"] = False
            job_status["error"] = None
            job_status["excel_path"] = None
            job_status["doc_count"] = 0

        def log(msg):
            with job_lock:
                job_status["log"].append(msg)

        log(f"Starte Export: {date_from} bis {date_to}")
        if tag_names:
            log(f"Tags: {', '.join(tag_names)}")

        # Dokumente abrufen
        log("Lade Dokumentenliste von Paperless...")
        docs = get_documents(date_from, date_to, tag_ids)
        log(f"{len(docs)} Dokumente gefunden.")

        if not docs:
            with job_lock:
                job_status["done"] = True
                job_status["running"] = False
                job_status["error"] = "Keine Dokumente im gewählten Zeitraum/Tag gefunden."
            return

        with job_lock:
            job_status["doc_count"] = len(docs)

        # Ausgabeordner – Schreibzugriff nur innerhalb OUTPUT_DIR
        export_folder = os.path.join(OUTPUT_DIR, year_label)
        pdf_folder = os.path.join(export_folder, "Belege")
        _assert_output_path(export_folder)  # Sicherheitsprüfung
        _assert_output_path(pdf_folder)
        os.makedirs(pdf_folder, exist_ok=True)
        log(f"Ausgabeordner: {export_folder}")

        # PDFs herunterladen
        log("Lade PDFs herunter...")
        pdf_map = download_pdfs(docs, pdf_folder, PAPERLESS_URL, PAPERLESS_TOKEN, log)

        # Excel erstellen
        log("Erstelle Excel-Datei...")
        excel_filename = f"Rechnungsaufstellung_{year_label}.xlsx"
        excel_path = os.path.join(export_folder, excel_filename)
        create_excel(docs, pdf_map, excel_path, year_label)
        log(f"Excel gespeichert: {excel_filename}")

        with job_lock:
            job_status["done"] = True
            job_status["running"] = False
            job_status["excel_path"] = excel_path
            job_status["log"].append(f"✓ Export abgeschlossen. {len(docs)} Belege exportiert.")

    except Exception as e:
        with job_lock:
            job_status["done"] = True
            job_status["running"] = False
            job_status["error"] = str(e)
            job_status["log"].append(f"✗ Fehler: {e}")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/tags")
def api_tags():
    try:
        tags = get_all_tags()
        return jsonify({"tags": tags})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/status")
def api_status():
    with job_lock:
        return jsonify(dict(job_status))


@app.route("/api/start", methods=["POST"])
def api_start():
    global job_status
    with job_lock:
        if job_status["running"]:
            return jsonify({"error": "Ein Job läuft bereits."}), 400

    data = request.json
    date_from = data.get("date_from")
    date_to = data.get("date_to")
    tag_ids = data.get("tag_ids", [])
    tag_names = data.get("tag_names", [])
    year_label = data.get("year_label", date_from[:4] if date_from else "export")

    if not date_from or not date_to:
        return jsonify({"error": "Bitte Datumsbereich angeben."}), 400

    with job_lock:
        job_status["running"] = True
        job_status["done"] = False
        job_status["log"] = []
        job_status["error"] = None

    thread = threading.Thread(
        target=run_export_job,
        args=(date_from, date_to, tag_ids, tag_names, year_label),
        daemon=True,
    )
    thread.start()
    return jsonify({"ok": True})


@app.route("/api/config")
def api_config():
    """Liefert unkritische Konfigurationsinfos für die UI."""
    return jsonify({
        "output_dir": OUTPUT_DIR,
        "paperless_url": PAPERLESS_URL,
    })


@app.route("/api/health")
def api_health():
    """Diagnose-Endpunkt: prüft ob Env-Vars gesetzt sind und Paperless erreichbar ist."""
    token_set = bool(PAPERLESS_TOKEN)
    status = {
        "paperless_url": PAPERLESS_URL,
        "token_configured": token_set,
        "output_dir": OUTPUT_DIR,
    }
    if not token_set:
        status["error"] = "PAPERLESS_TOKEN ist nicht gesetzt (leere Umgebungsvariable)."
        return jsonify(status), 200  # 200 damit Frontend es lesen kann
    try:
        data = paperless_get("/api/tags/", params={"page_size": 1})
        status["paperless_reachable"] = True
        status["tags_count"] = data.get("count", "?")
    except Exception as e:
        status["paperless_reachable"] = False
        status["error"] = str(e)
    return jsonify(status)


@app.route("/api/download-excel")
def api_download_excel():
    with job_lock:
        path = job_status.get("excel_path")
    if not path or not os.path.exists(path):
        return jsonify({"error": "Keine Excel-Datei verfügbar."}), 404
    # Sicherheitsprüfung: nur Dateien aus OUTPUT_DIR ausliefern
    try:
        _assert_output_path(path)
    except PermissionError:
        return jsonify({"error": "Zugriff verweigert."}), 403
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
