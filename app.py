import os
import re
import io
import json
import requests
import threading
import time as _time
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, send_file

from excel_export import create_excel, update_excel_with_ocr
from pdf_export import download_pdfs
from llm_extract import extract_from_ocr, check_ollama_available

app = Flask(__name__)

# Cache-Busting: Static-Files bekommen einen Versions-Query-Parameter
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
_BUILD_VERSION = str(int(_time.time()))

@app.context_processor
def inject_version():
    return {"ver": _BUILD_VERSION}

PAPERLESS_URL   = os.environ.get("PAPERLESS_URL",   "http://192.168.178.115:8000")
PAPERLESS_TOKEN = os.environ.get("PAPERLESS_TOKEN", "")
OUTPUT_DIR      = os.environ.get("OUTPUT_DIR",      "/output")
OLLAMA_URL      = os.environ.get("OLLAMA_URL",      "http://192.168.178.115:11434")
OLLAMA_MODEL    = os.environ.get("OLLAMA_MODEL",    "qwen2.5:3b")
WINDOWS_UNC_PATH = os.environ.get("WINDOWS_UNC_PATH", r"\\SynologyDS923\downloads\steuerberater")

# Ollama Env-Vars weitergeben an llm_extract Modul
os.environ.setdefault("OLLAMA_URL",   OLLAMA_URL)
os.environ.setdefault("OLLAMA_MODEL", OLLAMA_MODEL)

# ── Sicherheit ─────────────────────────────────────────────────────────
ALLOWED_PAPERLESS_METHODS = {"GET"}

def _assert_output_path(path: str):
    real_output = os.path.realpath(OUTPUT_DIR)
    real_path   = os.path.realpath(path)
    if not real_path.startswith(real_output):
        raise PermissionError(
            f"Schreibzugriff außerhalb des erlaubten Pfads verweigert: {path}"
        )

# ── Globaler Job-Status ────────────────────────────────────────────────
job_status = {
    "running":    False,
    "stage":      None,    # "stage1" | "stage2"
    "log":        [],
    "done":       False,
    "error":      None,
    "excel_path": None,
    "doc_count":  0,
    "ocr_current": 0,      # aktuelles Dokument Stufe 2
    "ocr_total":   0,      # Gesamtanzahl Stufe 2
    "ocr_current_title": "",
}
job_lock = threading.Lock()


def paperless_get(path, params=None):
    """Einziger HTTP-Zugang zu Paperless – ausschließlich GET (read-only)."""
    if not path.startswith("/api/"):
        raise ValueError(f"Ungültiger API-Pfad: {path}")
    headers = {"Authorization": f"Token {PAPERLESS_TOKEN}"}
    url  = f"{PAPERLESS_URL.rstrip('/')}{path}"
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_all_tags():
    tags = []
    url  = "/api/tags/"
    while url:
        data = paperless_get(url)
        tags.extend(data.get("results", []))
        next_url = data.get("next")
        if next_url:
            from urllib.parse import urlparse
            parsed = urlparse(next_url)
            url = parsed.path + ("?" + parsed.query if parsed.query else "")
            if not url.startswith("/api/"):
                url = None
        else:
            url = None
    return tags


def get_all_correspondents():
    """Holt alle Correspondents als {id: name} Dict."""
    result = {}
    url    = "/api/correspondents/"
    while url:
        data = paperless_get(url)
        for c in data.get("results", []):
            result[c["id"]] = c["name"]
        next_url = data.get("next")
        if next_url:
            from urllib.parse import urlparse
            parsed = urlparse(next_url)
            url = parsed.path + ("?" + parsed.query if parsed.query else "")
            if not url.startswith("/api/"):
                url = None
        else:
            url = None
    return result


def get_documents(date_from, date_to, tag_ids):
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


def enrich_documents_with_correspondents(documents, correspondents):
    """Hängt correspondent_name an jedes Dokument."""
    for doc in documents:
        cid = doc.get("correspondent")
        doc["correspondent_name"] = correspondents.get(cid) if cid else None
    return documents


# ── Stufe 1: Download + Excel ──────────────────────────────────────────
def run_stage1(date_from, date_to, tag_ids, tag_names, year_label):
    global job_status
    try:
        with job_lock:
            job_status.update({
                "log": [], "done": False, "error": None,
                "excel_path": None, "doc_count": 0, "stage": "stage1",
                "ocr_current": 0, "ocr_total": 0, "ocr_current_title": "",
            })

        def log(msg):
            with job_lock:
                job_status["log"].append(msg)

        log(f"Stufe 1 gestartet: {date_from} bis {date_to}")
        if tag_names:
            log(f"Tags: {', '.join(tag_names)}")

        log("Lade Correspondents aus Paperless…")
        correspondents = get_all_correspondents()
        log(f"{len(correspondents)} Correspondents geladen.")

        log("Lade Dokumentenliste…")
        docs = get_documents(date_from, date_to, tag_ids)
        docs = enrich_documents_with_correspondents(docs, correspondents)
        log(f"{len(docs)} Dokumente gefunden.")

        if not docs:
            with job_lock:
                job_status.update({"done": True, "running": False,
                    "error": "Keine Dokumente im gewählten Zeitraum/Tag gefunden."})
            return

        with job_lock:
            job_status["doc_count"] = len(docs)

        export_folder = os.path.join(OUTPUT_DIR, year_label)
        pdf_folder    = os.path.join(export_folder, "Belege")
        _assert_output_path(export_folder)
        _assert_output_path(pdf_folder)
        os.makedirs(pdf_folder, exist_ok=True)
        log(f"Ausgabeordner: {export_folder}")

        log("Lade PDFs herunter…")
        pdf_map = download_pdfs(docs, pdf_folder, PAPERLESS_URL, PAPERLESS_TOKEN, log)

        log("Erstelle Excel-Datei…")
        excel_filename = f"Rechnungsaufstellung_{year_label}.xlsx"
        excel_path     = os.path.join(export_folder, excel_filename)
        create_excel(docs, pdf_map, excel_path, year_label,
                     unc_base=WINDOWS_UNC_PATH)
        log(f"Excel gespeichert: {excel_filename}")

        with job_lock:
            job_status.update({
                "done": True, "running": False,
                "excel_path": excel_path,
                "log": job_status["log"] + [f"✓ Stufe 1 abgeschlossen. {len(docs)} Belege exportiert."],
            })

    except Exception as e:
        with job_lock:
            job_status.update({
                "done": True, "running": False, "error": str(e),
                "log": job_status["log"] + [f"✗ Fehler: {e}"],
            })


# ── Stufe 2: OCR-Analyse via Ollama ───────────────────────────────────
def run_stage2(excel_path, year_label, docs=None,
               date_from=None, date_to=None, tag_ids=None):
    """
    Liest PDFs aus dem bestehenden Export-Ordner,
    holt OCR-Text aus Paperless und schreibt Ergebnisse ins Excel.
    """
    global job_status
    try:
        with job_lock:
            job_status.update({
                "log": [], "done": False, "error": None,
                "stage": "stage2", "ocr_current": 0, "ocr_total": 0,
                "ocr_current_title": "",
            })

        def log(msg):
            with job_lock:
                job_status["log"].append(msg)

        # Ollama verfügbar?
        ok, msg = check_ollama_available()
        if not ok:
            raise RuntimeError(f"Ollama nicht verfügbar: {msg}")
        log(f"Ollama bereit: {msg}")

        # Dokumente laden falls nicht übergeben
        if docs is None:
            if not (date_from and date_to):
                raise ValueError("Kein Datumsbereich für Stufe 2 angegeben.")
            log("Lade Dokumentenliste…")
            correspondents = get_all_correspondents()
            docs = get_documents(date_from, date_to, tag_ids or [])
            docs = enrich_documents_with_correspondents(docs, correspondents)
            log(f"{len(docs)} Dokumente geladen.")

        with job_lock:
            job_status["ocr_total"] = len(docs)

        log(f"Starte OCR-Analyse für {len(docs)} Dokumente…")
        ocr_results = {}

        for idx, doc in enumerate(docs, start=1):
            doc_id = doc.get("id")
            title  = doc.get("title", f"Dokument {doc_id}")

            with job_lock:
                job_status["ocr_current"]       = idx
                job_status["ocr_current_title"] = title

            log(f"[{idx}/{len(docs)}] Analysiere: {title}")

            # Correspondent schon bekannt? Dann nur Betrag via LLM
            has_correspondent = bool(doc.get("correspondent_name"))

            content = doc.get("content", "")
            if not content:
                # Aus Paperless nachladen (content ist manchmal nicht im List-Endpoint)
                try:
                    detail = paperless_get(f"/api/documents/{doc_id}/")
                    content = detail.get("content", "")
                except Exception:
                    pass

            result = extract_from_ocr(content)

            if result.get("error"):
                log(f"  ⚠ LLM-Fehler: {result['error']}")

            ocr_results[doc_id] = {
                "absender": None if has_correspondent else result.get("absender"),
                "betrag":   result.get("betrag"),
            }

            if result.get("absender") and not has_correspondent:
                log(f"  Absender: {result['absender']}")
            if result.get("betrag") is not None:
                log(f"  Betrag: {result['betrag']:.2f} €")

        log("Schreibe OCR-Ergebnisse ins Excel…")
        updated = update_excel_with_ocr(
            excel_path, ocr_results, WINDOWS_UNC_PATH, year_label
        )
        log(f"✓ Stufe 2 abgeschlossen. {updated} Felder aktualisiert.")

        with job_lock:
            job_status.update({"done": True, "running": False})

    except Exception as e:
        with job_lock:
            job_status.update({
                "done": True, "running": False, "error": str(e),
                "log": job_status["log"] + [f"✗ Fehler: {e}"],
            })


# ── Flask Routes ───────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/tags")
def api_tags():
    try:
        return jsonify({"tags": get_all_tags()})
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

    data       = request.json
    date_from  = data.get("date_from")
    date_to    = data.get("date_to")
    tag_ids    = data.get("tag_ids", [])
    tag_names  = data.get("tag_names", [])
    year_label = data.get("year_label", date_from[:4] if date_from else "export")
    mode       = data.get("mode", "stage1")  # "stage1" | "stage2" | "both"

    if not date_from or not date_to:
        return jsonify({"error": "Bitte Datumsbereich angeben."}), 400

    with job_lock:
        job_status.update({"running": True, "done": False,
                           "log": [], "error": None})

    if mode == "stage2":
        # Stufe 2 standalone: Excel muss bereits existieren
        excel_path = os.path.join(OUTPUT_DIR, year_label,
                                  f"Rechnungsaufstellung_{year_label}.xlsx")
        if not os.path.exists(excel_path):
            with job_lock:
                job_status["running"] = False
            return jsonify({
                "error": f"Excel nicht gefunden. Bitte zuerst Stufe 1 ausführen."
            }), 400
        with job_lock:
            job_status["excel_path"] = excel_path
        thread = threading.Thread(
            target=run_stage2,
            args=(excel_path, year_label, None, date_from, date_to, tag_ids),
            daemon=True,
        )
    elif mode == "both":
        # Beide Stufen sequenziell in einem Thread
        def run_both():
            run_stage1(date_from, date_to, tag_ids, tag_names, year_label)
            with job_lock:
                ep = job_status.get("excel_path")
                err = job_status.get("error")
            if ep and not err:
                with job_lock:
                    job_status.update({"running": True, "done": False})
                run_stage2(ep, year_label, None, date_from, date_to, tag_ids)
        thread = threading.Thread(target=run_both, daemon=True)
    else:
        # Stufe 1 only
        thread = threading.Thread(
            target=run_stage1,
            args=(date_from, date_to, tag_ids, tag_names, year_label),
            daemon=True,
        )

    thread.start()
    return jsonify({"ok": True})


@app.route("/api/config")
def api_config():
    return jsonify({
        "output_dir":       OUTPUT_DIR,
        "paperless_url":    PAPERLESS_URL,
        "windows_unc_path": WINDOWS_UNC_PATH,
        "ollama_model":     OLLAMA_MODEL,
    })


@app.route("/api/health")
def api_health():
    token_set = bool(PAPERLESS_TOKEN)
    status = {
        "paperless_url":    PAPERLESS_URL,
        "token_configured": token_set,
        "output_dir":       OUTPUT_DIR,
        "ollama_url":       OLLAMA_URL,
        "ollama_model":     OLLAMA_MODEL,
    }
    if not token_set:
        status["error"] = "PAPERLESS_TOKEN ist nicht gesetzt."
        return jsonify(status)
    try:
        data = paperless_get("/api/tags/", params={"page_size": 1})
        status["paperless_reachable"] = True
        status["tags_count"] = data.get("count", "?")
    except Exception as e:
        status["paperless_reachable"] = False
        status["error"] = str(e)

    # Ollama Check
    ok, msg = check_ollama_available()
    status["ollama_available"] = ok
    status["ollama_status"]    = msg

    return jsonify(status)


@app.route("/api/download-excel")
def api_download_excel():
    with job_lock:
        path = job_status.get("excel_path")
    if not path or not os.path.exists(path):
        return jsonify({"error": "Keine Excel-Datei verfügbar."}), 404
    try:
        _assert_output_path(path)
    except PermissionError:
        return jsonify({"error": "Zugriff verweigert."}), 403
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
