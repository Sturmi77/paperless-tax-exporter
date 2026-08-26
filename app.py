import os
import re
import io
import json
import requests
import threading
import time as _time
from datetime import datetime, date
from urllib.parse import urlparse
from flask import Flask, render_template, request, jsonify, send_file

from excel_export import create_excel, update_excel_with_ocr, \
                         append_to_excel, get_existing_doc_ids, \
                         read_invoices_for_matching, update_excel_with_bank_matches, \
                         prepare_stb_export
from pdf_export import download_pdfs
from llm_extract import extract_from_ocr, check_ollama_available
from bank_csv import parse_bank_csv, csv_preview
from matching_csv import match_invoices_to_bank

app = Flask(__name__)

# Cache-Busting: Static-Files bekommen einen Versions-Query-Parameter
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
_BUILD_VERSION = str(int(_time.time()))

@app.context_processor
def inject_version():
    return {"ver": _BUILD_VERSION}

PAPERLESS_URL    = os.environ.get("PAPERLESS_URL",    "http://192.168.178.115:8000")
PAPERLESS_TOKEN  = os.environ.get("PAPERLESS_TOKEN",  "")
OUTPUT_DIR       = os.path.realpath(os.environ.get("OUTPUT_DIR", "/output"))
OLLAMA_URL       = os.environ.get("OLLAMA_URL",       "http://192.168.178.115:11434")
OLLAMA_MODEL     = os.environ.get("OLLAMA_MODEL",     "qwen2.5:3b")
WINDOWS_UNC_PATH = os.environ.get("WINDOWS_UNC_PATH", r"\\SynologyDS923\downloads\steuerberater")
# Issue #8: Hyperlink-Modus
# "cell"  = CELL("filename")-Formel → portabel, funktioniert nach Ordner-Verschiebung
# "unc"   = absoluter UNC-Pfad (bisheriges Verhalten, Fallback)
HYPERLINK_MODE    = os.environ.get("HYPERLINK_MODE", "cell").lower()
# Issue #8: Spalte K mit kopierbarem UNC-Pfad (plain text) einblenden
INCLUDE_TEXT_PATH = os.environ.get("INCLUDE_TEXT_PATH", "false").lower() == "true"
APP_TITLE         = os.environ.get("APP_TITLE", "Steuerberater Export")

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
    "stage":      None,       # "stage0" | "stage1" | "stage2"
    "log":        [],
    "done":       False,
    "error":      None,
    "cancelled":  False,      # Issue #2
    "excel_path": None,
    "doc_count":  0,
    # Stage 0 (Excel) – Issue #19
    "s0_current": 0,
    "s0_total":   0,
    "s0_current_title": "",
    "s0_start_time":    None,
    "s0_last_doc_time": None,
    # Stage 1 (PDF-Download) – Issue #19
    "s1_current": 0,
    "s1_total":   0,
    "s1_current_title": "",
    "s1_start_time":    None,
    "s1_last_doc_time": None,
    # Stage 2 (OCR)
    "ocr_current": 0,
    "ocr_total":   0,
    "ocr_current_title": "",
    "ocr_start_time":    None,
    "ocr_last_doc_time": None,
    "cancellable":       False,    # True sobald Stufe 2 gestartet, bis done
}
job_lock     = threading.Lock()
cancel_event = threading.Event()   # Issue #2: Abbruch-Steuerung


def _calc_eta(start_time, last_doc_time, done, total):
    """
    Ø-Zeit und ETA in Sekunden (Issue #19).
    Rückgabe: (avg_sec_per_doc, eta_seconds) oder (None, None).
    """
    if not start_time or done == 0 or total == 0:
        return None, None
    elapsed = (last_doc_time or _time.monotonic()) - start_time
    avg = round(elapsed / done, 1)
    eta = round(avg * (total - done))
    return avg, eta


def _validate_subfolder(name: str) -> str:
    """
    Allowlist-Validierung für Unterordner-Namen (Issue #9).
    Erlaubt: Buchstaben, Zahlen, Unterstriche, Bindestriche (1-50 Zeichen).
    Leerer String → kein Unterordner (akzeptiert).
    """
    if not name:
        return ""
    name = name.strip()
    if not name:
        return ""
    if not re.fullmatch(r"[A-Za-z0-9_\-]{1,50}", name):
        raise ValueError(
            f"Ungültiger Unterordner-Name: '{name}'. "
            "Nur Buchstaben, Zahlen, _ und - erlaubt (max. 50 Zeichen)."
        )
    return name


def _job_status_reset(stage):
    """Job-Status für neuen Lauf initialisieren."""
    is_stage2 = (stage == "stage2")
    with job_lock:
        job_status.update({
            "log": [], "done": False, "error": None, "cancelled": False,
            "excel_path": None, "doc_count": 0, "stage": stage,
            "s0_current": 0, "s0_total": 0, "s0_current_title": "",
            "s0_start_time": None, "s0_last_doc_time": None,
            "s1_current": 0, "s1_total": 0, "s1_current_title": "",
            "s1_start_time": None, "s1_last_doc_time": None,
            "ocr_current": 0, "ocr_total": 0, "ocr_current_title": "",
            "ocr_start_time": None, "ocr_last_doc_time": None,
            "cancellable": is_stage2,  # sofort ab Start von stage2 abbrechbar
        })
    cancel_event.clear()


def _log(msg):
    with job_lock:
        job_status["log"].append(msg)


# ── Paperless API (read-only) ──────────────────────────────────────────
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


def get_all_document_types():
    """Holt alle Document Types als Liste von {id, name}."""
    types = []
    url   = "/api/document_types/"
    while url:
        data = paperless_get(url)
        for t in data.get("results", []):
            types.append({"id": t["id"], "name": t["name"]})
        next_url = data.get("next")
        if next_url:
            from urllib.parse import urlparse
            parsed = urlparse(next_url)
            url = parsed.path + ("?" + parsed.query if parsed.query else "")
            if not url.startswith("/api/"):
                url = None
        else:
            url = None
    return types


def get_documents(date_from, date_to, tag_ids, date_field="created",
                  document_type_ids=None):
    """Issue #3: date_field = 'created' (Belegdatum) oder 'added' (Scan-Datum).

    Paginierung folgt der 'next'-URL von Paperless direkt, damit Cursor-basierte
    Paginierung und alle Filter-Parameter auf jeder Seite erhalten bleiben.
    """
    documents = []
    params = {
        f"{date_field}__date__gte": date_from,
        f"{date_field}__date__lte": date_to,
        "page_size": 100,
        "ordering": date_field,  # stabile Sortierung für konsistente Paginierung
    }
    if tag_ids:
        params["tags__id__all"] = ",".join(str(t) for t in tag_ids)
    if document_type_ids:
        params["document_type__id__in"] = ",".join(str(t) for t in document_type_ids)

    # Erste Seite mit vollen params abrufen
    data = paperless_get("/api/documents/", params=params)
    documents.extend(data.get("results", []))

    # Folgeseiten via next-URL – Filter-Parameter sind bereits in der URL enthalten
    while data.get("next"):
        parsed = urlparse(data["next"])
        next_path = parsed.path + ("?" + parsed.query if parsed.query else "")
        if not next_path.startswith("/api/"):
            break
        data = paperless_get(next_path)  # keine eigenen params – URL enthält alles
        documents.extend(data.get("results", []))

    return documents


def enrich_documents_with_correspondents(documents, correspondents):
    for doc in documents:
        cid = doc.get("correspondent")
        doc["correspondent_name"] = correspondents.get(cid) if cid else None
    return documents


# ── Stufe 0: Nur Excel (Issue #1) ─────────────────────────────────────
def run_stage0(date_from, date_to, tag_ids, tag_names, year_label, date_field="created",
               document_type_ids=None, subfolder: str = ""):
    """Erstellt nur die Excel-Datei – kein PDF-Download, kein OCR."""
    global job_status
    try:
        _job_status_reset("stage0")
        _log(f"Nur Excel – gestartet: {date_from} bis {date_to}")
        if tag_names:
            _log(f"Tags: {', '.join(tag_names)}")
        _log(f"Datumsfeld: {'Scan-Datum' if date_field == 'added' else 'Belegdatum'}")
        if document_type_ids:
            _log(f"Dokumenttyp-Filter: {len(document_type_ids)} Typ(en)")
        if subfolder:
            _log(f"Unterordner: {subfolder}")

        _log("Lade Correspondents aus Paperless…")
        correspondents = get_all_correspondents()
        _log(f"{len(correspondents)} Correspondents geladen.")

        _log("Lade Dokumentenliste…")
        docs = get_documents(date_from, date_to, tag_ids, date_field, document_type_ids)
        docs = enrich_documents_with_correspondents(docs, correspondents)
        _log(f"{len(docs)} Dokumente gefunden.")

        if not docs:
            with job_lock:
                job_status.update({"done": True, "running": False,
                    "error": "Keine Dokumente im gewählten Zeitraum/Tag gefunden."})
            return

        with job_lock:
            job_status["doc_count"] = len(docs)
            job_status["s0_total"] = len(docs)
            job_status["s0_start_time"] = _time.monotonic()

        export_folder = os.path.join(OUTPUT_DIR, year_label)
        _assert_output_path(export_folder)
        os.makedirs(export_folder, exist_ok=True)
        _log(f"Ausgabeordner: {export_folder}")

        def _s0_progress(idx, total, title):
            with job_lock:
                job_status["s0_current"] = idx
                job_status["s0_current_title"] = title or ""
                job_status["s0_last_doc_time"] = _time.monotonic()

        _log("Erstelle Excel-Datei…")
        excel_filename = f"Rechnungsaufstellung_{year_label}.xlsx"
        excel_path     = os.path.join(export_folder, excel_filename)
        create_excel(docs, {}, excel_path, year_label, unc_base=WINDOWS_UNC_PATH,
                      subfolder=subfolder, hyperlink_mode=HYPERLINK_MODE,
                      include_text_path=INCLUDE_TEXT_PATH,
                      progress_fn=_s0_progress)
        _log(f"Excel gespeichert: {excel_filename}")

        with job_lock:
            job_status.update({
                "done": True, "running": False, "excel_path": excel_path,
                "log": job_status["log"] + [f"✓ Excel abgeschlossen. {len(docs)} Belege exportiert."],
            })

    except Exception as e:
        with job_lock:
            job_status.update({
                "done": True, "running": False, "error": str(e),
                "log": job_status["log"] + [f"✗ Fehler: {e}"],
            })


# ── Stufe 1: Download + Excel ──────────────────────────────────────────
def run_stage1(date_from, date_to, tag_ids, tag_names, year_label,
               date_field="created", append_mode=False, document_type_ids=None,
               subfolder: str = ""):
    global job_status
    try:
        _job_status_reset("stage1")
        mode_label = "Nachtrag" if append_mode else "Stufe 1"
        _log(f"{mode_label} gestartet: {date_from} bis {date_to}")
        if tag_names:
            _log(f"Tags: {', '.join(tag_names)}")
        _log(f"Datumsfeld: {'Scan-Datum' if date_field == 'added' else 'Belegdatum'}")
        if subfolder:
            _log(f"Unterordner: {subfolder}")

        _log("Lade Correspondents aus Paperless…")
        correspondents = get_all_correspondents()
        _log(f"{len(correspondents)} Correspondents geladen.")

        _log("Lade Dokumentenliste…")
        docs = get_documents(date_from, date_to, tag_ids, date_field, document_type_ids)
        docs = enrich_documents_with_correspondents(docs, correspondents)
        _log(f"{len(docs)} Dokumente gefunden.")

        if not docs:
            with job_lock:
                job_status.update({"done": True, "running": False,
                    "error": "Keine Dokumente im gewählten Zeitraum/Tag gefunden."})
            return

        export_folder  = os.path.join(OUTPUT_DIR, year_label)
        # Subfolder: <year>/<subfolder>/Belege/ wenn gesetzt, sonst <year>/Belege/
        if subfolder:
            pdf_folder = os.path.join(export_folder, subfolder, "Belege")
        else:
            pdf_folder = os.path.join(export_folder, "Belege")
        excel_filename = f"Rechnungsaufstellung_{year_label}.xlsx"
        excel_path     = os.path.join(export_folder, excel_filename)
        _assert_output_path(export_folder)
        _assert_output_path(pdf_folder)
        os.makedirs(pdf_folder, exist_ok=True)
        _log(f"Ausgabeordner: {export_folder}")

        # Append-Modus: nur wirklich neue Dokumente verarbeiten
        if append_mode and os.path.exists(excel_path):
            existing_ids = get_existing_doc_ids(excel_path)
            _log(f"{len(existing_ids)} Belege bereits im Excel vorhanden.")
            new_docs = [
                d for d in docs
                if str(d.get("archive_serial_number") or d.get("id")) not in existing_ids
            ]
            _log(f"{len(new_docs)} neue Belege werden hinzugefügt.")
            if not new_docs:
                with job_lock:
                    job_status.update({"done": True, "running": False,
                        "error": "Keine neuen Dokumente gefunden – alle bereits im Excel vorhanden."})
                return
            docs_to_process = new_docs
        else:
            docs_to_process = docs

        with job_lock:
            job_status["doc_count"] = len(docs_to_process)
            job_status["s1_total"] = len(docs_to_process)
            job_status["s1_start_time"] = _time.monotonic()

        def _s1_progress(idx, total, title):
            with job_lock:
                job_status["s1_current"] = idx
                job_status["s1_current_title"] = title or ""
                job_status["s1_last_doc_time"] = _time.monotonic()

        _log(f"Lade {len(docs_to_process)} PDFs herunter…")
        pdf_map = download_pdfs(
            docs_to_process, pdf_folder, PAPERLESS_URL, PAPERLESS_TOKEN, _log,
            progress_fn=_s1_progress,
        )

        if append_mode and os.path.exists(excel_path):
            _log("Hänge neue Zeilen ans Excel an…")
            added = append_to_excel(docs_to_process, pdf_map, excel_path, year_label,
                                    unc_base=WINDOWS_UNC_PATH, subfolder=subfolder,
                                    hyperlink_mode=HYPERLINK_MODE,
                                    include_text_path=INCLUDE_TEXT_PATH)
            _log(f"Excel ergänzt: {added} neue Zeile(n) hinzugefügt.")
        else:
            _log("Erstelle Excel-Datei…")
            create_excel(docs_to_process, pdf_map, excel_path, year_label,
                         unc_base=WINDOWS_UNC_PATH, subfolder=subfolder,
                         hyperlink_mode=HYPERLINK_MODE,
                         include_text_path=INCLUDE_TEXT_PATH)
            _log(f"Excel gespeichert: {excel_filename}")

        with job_lock:
            job_status.update({
                "done": True, "running": False, "excel_path": excel_path,
                "log": job_status["log"] + [
                    f"✓ {mode_label} abgeschlossen. {len(docs_to_process)} Belege exportiert."
                ],
            })

    except Exception as e:
        with job_lock:
            job_status.update({
                "done": True, "running": False, "error": str(e),
                "log": job_status["log"] + [f"✗ Fehler: {e}"],
            })


# ── Stufe 2: OCR-Analyse via Ollama ───────────────────────────────────
def run_stage2(excel_path, year_label, docs=None,
               date_from=None, date_to=None, tag_ids=None, date_field="created",
               document_type_ids=None, subfolder: str = ""):
    """Issue #2: cancel_event wird nach jedem Dokument geprüft."""
    global job_status
    try:
        _job_status_reset("stage2")

        ok, msg = check_ollama_available()
        if not ok:
            raise RuntimeError(f"Ollama nicht verfügbar: {msg}")
        _log(f"Ollama bereit: {msg}")

        if docs is None:
            if not (date_from and date_to):
                raise ValueError("Kein Datumsbereich für Stufe 2 angegeben.")
            _log("Lade Dokumentenliste…")
            correspondents = get_all_correspondents()
            # Abbruch-Check nach Correspondents-Laden
            if cancel_event.is_set():
                _log("⚠ Job vor OCR-Start abgebrochen.")
                with job_lock:
                    job_status.update({"done": True, "running": False,
                                       "cancelled": True, "cancellable": False})
                return
            docs = get_documents(date_from, date_to, tag_ids or [], date_field, document_type_ids)
            docs = enrich_documents_with_correspondents(docs, correspondents)
            _log(f"{len(docs)} Dokumente geladen.")
            # Abbruch-Check nach Dokumente-Laden
            if cancel_event.is_set():
                _log("⚠ Job vor OCR-Start abgebrochen.")
                with job_lock:
                    job_status.update({"done": True, "running": False,
                                       "cancelled": True, "cancellable": False})
                return

        with job_lock:
            job_status["ocr_total"]      = len(docs)
            job_status["ocr_start_time"] = _time.monotonic()

        _log(f"Starte OCR-Analyse für {len(docs)} Dokumente…")
        ocr_results = {}

        for idx, doc in enumerate(docs, start=1):
            # Issue #2: Abbruch-Check
            if cancel_event.is_set():
                _log(f"⚠ Job abgebrochen nach {idx - 1} von {len(docs)} Dokumenten.")
                break

            doc_id = doc.get("id")
            title  = doc.get("title", f"Dokument {doc_id}")

            with job_lock:
                job_status["ocr_current"]       = idx
                job_status["ocr_current_title"] = title
                job_status["ocr_last_doc_time"] = _time.monotonic()

            _log(f"[{idx}/{len(docs)}] Analysiere: {title}")

            has_correspondent = bool(doc.get("correspondent_name"))
            content = doc.get("content", "")
            if not content:
                try:
                    detail  = paperless_get(f"/api/documents/{doc_id}/")
                    content = detail.get("content", "")
                except Exception:
                    pass

            result = extract_from_ocr(content)
            if result.get("error"):
                _log(f"  ⚠ LLM-Fehler: {result['error']}")

            ocr_results[doc_id] = {
                "absender": None if has_correspondent else result.get("absender"),
                "betrag":   result.get("betrag"),
            }
            if result.get("absender") and not has_correspondent:
                _log(f"  Absender: {result['absender']}")
            if result.get("betrag") is not None:
                _log(f"  Betrag: {result['betrag']:.2f} €")

        # Auch bei Abbruch: bisherige Ergebnisse sichern
        _log("Schreibe OCR-Ergebnisse ins Excel…")
        updated   = update_excel_with_ocr(excel_path, ocr_results, WINDOWS_UNC_PATH, year_label,
                                          subfolder=subfolder)
        cancelled = cancel_event.is_set()
        suffix    = " (abgebrochen)" if cancelled else ""
        _log(f"✓ Stufe 2 abgeschlossen. {updated} Felder aktualisiert.{suffix}")

        with job_lock:
            job_status.update({"done": True, "running": False,
                               "cancelled": cancelled, "cancellable": False})

    except Exception as e:
        with job_lock:
            job_status.update({
                "done": True, "running": False, "cancellable": False, "error": str(e),
                "log": job_status["log"] + [f"✗ Fehler: {e}"],
            })


# ── Flask Routes ───────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", app_title=APP_TITLE)


@app.route("/api/tags")
def api_tags():
    try:
        return jsonify({"tags": get_all_tags()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/document-types")
def api_document_types():
    try:
        return jsonify({"document_types": get_all_document_types()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/status")
def api_status():
    with job_lock:
        s = dict(job_status)

    # Issue #19: ETA für alle Stages über gemeinsame Hilfsfunktion
    s0_avg, s0_eta = _calc_eta(
        s.get("s0_start_time"), s.get("s0_last_doc_time"),
        s.get("s0_current", 0), s.get("s0_total", 0),
    )
    s1_avg, s1_eta = _calc_eta(
        s.get("s1_start_time"), s.get("s1_last_doc_time"),
        s.get("s1_current", 0), s.get("s1_total", 0),
    )
    ocr_avg, ocr_eta = _calc_eta(
        s.get("ocr_start_time"), s.get("ocr_last_doc_time"),
        s.get("ocr_current", 0), s.get("ocr_total", 0),
    )

    s["s0_avg_sec_per_doc"] = s0_avg
    s["s0_eta_seconds"]     = s0_eta
    s["s1_avg_sec_per_doc"] = s1_avg
    s["s1_eta_seconds"]     = s1_eta
    # Stage-2-Felder unverändert für bestehendes Frontend
    s["avg_sec_per_doc"] = ocr_avg
    s["eta_seconds"]     = ocr_eta

    for key in (
        "s0_start_time", "s0_last_doc_time",
        "s1_start_time", "s1_last_doc_time",
        "ocr_start_time", "ocr_last_doc_time",
    ):
        s.pop(key, None)
    return jsonify(s)


@app.route("/api/cancel", methods=["POST"])
def api_cancel():
    """Issue #2: Laufenden Stufe-2-Job abbrechen (auch während Ladevorgang)."""
    with job_lock:
        cancellable = job_status.get("cancellable", False)
    if not cancellable:
        return jsonify({"error": "Kein abbrechbarer Job aktiv."}), 400
    cancel_event.set()
    _log("⚠ Abbruch angefordert – wird nach aktuellem Schritt gestoppt…")
    return jsonify({"ok": True})


@app.route("/api/check-exists")
def api_check_exists():
    """Issue #4: Prüft ob Export-Dateien für ein Jahr bereits vorhanden sind."""
    year = request.args.get("year", "")
    if not year:
        return jsonify({"error": "Parameter 'year' fehlt."}), 400

    export_folder  = os.path.join(OUTPUT_DIR, year)
    excel_path     = os.path.join(export_folder, f"Rechnungsaufstellung_{year}.xlsx")
    pdf_folder     = os.path.join(export_folder, "Belege")

    excel_exists = os.path.isfile(excel_path)
    pdfs_exist   = os.path.isdir(pdf_folder)
    pdf_count    = len([f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]) if pdfs_exist else 0

    return jsonify({
        "excel_exists": excel_exists,
        "pdfs_exist":   pdfs_exist,
        "pdf_count":    pdf_count,
        "excel_name":   f"Rechnungsaufstellung_{year}.xlsx" if excel_exists else None,
    })


@app.route("/api/start", methods=["POST"])
def api_start():
    global job_status
    with job_lock:
        if job_status["running"]:
            return jsonify({"error": "Ein Job läuft bereits."}), 400

    data                = request.json
    date_from           = data.get("date_from")
    date_to             = data.get("date_to")
    tag_ids             = data.get("tag_ids", [])
    tag_names           = data.get("tag_names", [])
    year_label          = data.get("year_label", date_from[:4] if date_from else "export")
    mode                = data.get("mode", "stage1")   # "stage0"|"stage1"|"stage2"|"both"
    date_field          = data.get("date_field", "created")  # Issue #3
    append_mode         = data.get("append_mode", False)      # Issue #4: nur neue hinzufügen
    document_type_ids   = data.get("document_type_ids", []) or []  # Issue #7
    subfolder_raw       = data.get("subfolder", "") or ""              # Issue #9
    try:
        subfolder = _validate_subfolder(subfolder_raw)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if not date_from or not date_to:
        return jsonify({"error": "Bitte Datumsbereich angeben."}), 400

    with job_lock:
        job_status.update({"running": True, "done": False, "log": [], "error": None})

    if mode == "stage0":
        thread = threading.Thread(
            target=run_stage0,
            args=(date_from, date_to, tag_ids, tag_names, year_label, date_field),
            kwargs={"document_type_ids": document_type_ids, "subfolder": subfolder},
            daemon=True,
        )
    elif mode == "stage2":
        excel_path = os.path.join(OUTPUT_DIR, year_label,
                                  f"Rechnungsaufstellung_{year_label}.xlsx")
        if not os.path.exists(excel_path):
            with job_lock:
                job_status["running"] = False
            return jsonify({
                "error": "Excel nicht gefunden. Bitte zuerst Stufe 1 ausführen."
            }), 400
        with job_lock:
            job_status["excel_path"] = excel_path
        thread = threading.Thread(
            target=run_stage2,
            args=(excel_path, year_label, None, date_from, date_to, tag_ids, date_field),
            kwargs={"document_type_ids": document_type_ids, "subfolder": subfolder},
            daemon=True,
        )
    elif mode == "both":
        def run_both():
            run_stage1(date_from, date_to, tag_ids, tag_names, year_label, date_field,
                       append_mode=append_mode, document_type_ids=document_type_ids,
                       subfolder=subfolder)
            with job_lock:
                ep  = job_status.get("excel_path")
                err = job_status.get("error")
            if ep and not err:
                with job_lock:
                    job_status.update({"running": True, "done": False})
                cancel_event.clear()  # Sicherstellen dass kein alter Abbruch-State hängt
                run_stage2(ep, year_label, None, date_from, date_to, tag_ids, date_field,
                           document_type_ids=document_type_ids, subfolder=subfolder)
        thread = threading.Thread(target=run_both, daemon=True)
    else:
        # stage1
        thread = threading.Thread(
            target=run_stage1,
            args=(date_from, date_to, tag_ids, tag_names, year_label, date_field),
            kwargs={"append_mode": append_mode, "document_type_ids": document_type_ids,
                    "subfolder": subfolder},
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

    ok, msg = check_ollama_available()
    status["ollama_available"] = ok
    status["ollama_status"]    = msg

    return jsonify(status)


# ── Subfolder-Picker API (Issue #12, Schritt 4) ──────────────────────────────
@app.route("/api/subfolders")
def api_subfolders_list():
    """
    GET /api/subfolders
    Gibt alle direkten Unterverzeichnisse von OUTPUT_DIR zurueck (alphabetisch).
    Keine Rekursion, nur eine Ebene tief.
    """
    try:
        real_output = os.path.realpath(OUTPUT_DIR)
        if not os.path.isdir(real_output):
            return jsonify({"subfolders": []})
        names = sorted(
            entry.name
            for entry in os.scandir(real_output)
            if entry.is_dir(follow_symlinks=False)
            and re.fullmatch(r"[A-Za-z0-9_\-]{1,50}", entry.name)
        )
        return jsonify({"subfolders": names})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/subfolders", methods=["POST"])
def api_subfolders_create():
    """
    POST /api/subfolders  {"name": "<ordnername>"}
    Legt einen neuen Unterordner in OUTPUT_DIR an.
    Validiert via Allowlist (_validate_subfolder) + _assert_output_path.
    """
    data = request.get_json(force=True, silent=True) or {}
    raw_name = data.get("name", "") or ""
    try:
        name = _validate_subfolder(raw_name)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    if not name:
        return jsonify({"error": "Ordnername darf nicht leer sein."}), 400
    try:
        target = os.path.join(OUTPUT_DIR, name)
        _assert_output_path(target)
        os.makedirs(target, exist_ok=True)
        return jsonify({"created": name}), 201
    except PermissionError as e:
        return jsonify({"error": str(e)}), 403
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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


# ── Issue #25 / #28 / #29: Kontoauszug-Matching + Dateiauswahl ─────────

_PATH_SEG_RE = re.compile(r"^[A-Za-z0-9_\.\-]{1,120}$")


def _excel_path_for_year(year_label: str, subfolder: str = "") -> str:
    """Rechnungs-Excel liegt unter OUTPUT_DIR/{year}/ (Subfolder betrifft nur PDFs)."""
    year_label = re.sub(r"[^\w\-]", "", year_label or "")
    if not year_label:
        raise ValueError("Jahr fehlt.")
    path = os.path.join(OUTPUT_DIR, year_label, f"Rechnungsaufstellung_{year_label}.xlsx")
    _assert_output_path(path)
    return path


def _normalize_rel_path(rel: str) -> str:
    """Relativen Pfad unter OUTPUT_DIR validieren (kein Path-Traversal)."""
    if not rel or not str(rel).strip():
        raise ValueError("Relativpfad fehlt.")
    rel = str(rel).replace("\\", "/").strip().lstrip("/")
    parts = [p for p in rel.split("/") if p and p != "."]
    if not parts or any(p == ".." for p in parts):
        raise ValueError("Ungültiger Relativpfad.")
    for p in parts:
        if not _PATH_SEG_RE.fullmatch(p):
            raise ValueError(f"Ungültiges Pfadsegment: {p}")
    return "/".join(parts)


def _resolve_output_rel(rel: str) -> str:
    """Relativpfad → absoluter Pfad unter OUTPUT_DIR (realpath + Allowlist)."""
    norm = _normalize_rel_path(rel)
    abs_path = os.path.realpath(os.path.join(OUTPUT_DIR, *norm.split("/")))
    _assert_output_path(abs_path)
    return abs_path


def _list_output_files(extensions: tuple, max_depth: int = 3) -> list:
    """
    Listet Dateien unter OUTPUT_DIR mit gegebener Extension (klein, mit Punkt).
    max_depth: 0 = nur OUTPUT_DIR selbst.
    """
    exts = tuple(e.lower() if e.startswith(".") else f".{e.lower()}" for e in extensions)
    real_output = os.path.realpath(OUTPUT_DIR)
    results = []
    if not os.path.isdir(real_output):
        return results

    def walk(dir_path: str, depth: int, rel_parts: list):
        try:
            entries = list(os.scandir(dir_path))
        except OSError:
            return
        for entry in entries:
            name = entry.name
            if name.startswith(".") and name not in ("_uploads",):
                continue
            if not _PATH_SEG_RE.fullmatch(name):
                continue
            try:
                is_dir = entry.is_dir(follow_symlinks=False)
                is_file = entry.is_file(follow_symlinks=False)
            except OSError:
                continue
            rel = "/".join(rel_parts + [name])
            if is_dir and depth < max_depth:
                walk(entry.path, depth + 1, rel_parts + [name])
            elif is_file and name.lower().endswith(exts):
                try:
                    st = entry.stat(follow_symlinks=False)
                    mtime = int(st.st_mtime)
                    size = int(st.st_size)
                except OSError:
                    mtime, size = 0, 0
                results.append({
                    "rel_path": rel,
                    "name": name,
                    "mtime": mtime,
                    "size": size,
                })

    walk(real_output, 0, [])
    results.sort(key=lambda x: (-x["mtime"], x["rel_path"].lower()))
    return results


def _save_uploaded_file(file_storage, prefix: str, allowed_ext: tuple) -> str:
    """Speichert Upload unter OUTPUT_DIR/_uploads/ (Allowlist-Dateiname)."""
    upload_dir = os.path.join(OUTPUT_DIR, "_uploads")
    os.makedirs(upload_dir, exist_ok=True)
    _assert_output_path(upload_dir)
    orig = file_storage.filename or f"upload{allowed_ext[0]}"
    base = os.path.basename(orig)
    safe = re.sub(r"[^A-Za-z0-9_\.\-]", "_", base)[:120]
    lower = safe.lower()
    if not any(lower.endswith(ext) for ext in allowed_ext):
        safe += allowed_ext[0]
    prefix = re.sub(r"[^A-Za-z0-9_\-]", "", prefix or "file")[:40] or "file"
    target = os.path.join(upload_dir, f"{prefix}_{safe}")
    _assert_output_path(target)
    file_storage.save(target)
    rel = f"_uploads/{os.path.basename(target)}"
    return target, rel


def _save_uploaded_csv(file_storage, year_label: str) -> str:
    path, _rel = _save_uploaded_file(file_storage, year_label or "csv", (".csv",))
    return path


def _resolve_csv_path(data: dict, file_storage=None) -> tuple:
    """
    CSV aus Upload, csv_rel_path oder upload_key.
    Rückgabe: (abs_path, rel_path_or_key)
    """
    if file_storage and getattr(file_storage, "filename", None):
        path, rel = _save_uploaded_file(
            file_storage, data.get("year_label") or "csv", (".csv",)
        )
        return path, rel

    rel = (data.get("csv_rel_path") or "").strip()
    if rel:
        path = _resolve_output_rel(rel)
        if not os.path.isfile(path) or not path.lower().endswith(".csv"):
            raise FileNotFoundError("CSV nicht gefunden oder ungültig.")
        return path, _normalize_rel_path(rel)

    key = data.get("upload_key") or data.get("saved_path") or ""
    key = os.path.basename(str(key))
    if key and re.fullmatch(r"[A-Za-z0-9_\.\-]{1,140}", key):
        path = os.path.join(OUTPUT_DIR, "_uploads", key)
        _assert_output_path(path)
        if not os.path.isfile(path):
            raise FileNotFoundError("Hochgeladene CSV nicht gefunden.")
        return path, f"_uploads/{key}"

    raise ValueError("Keine CSV angegeben (Auswahl oder Upload).")


def _resolve_excel_path(data: dict, file_storage=None) -> tuple:
    """
    Excel aus excel_rel_path, Upload oder year_label-Fallback.
    Rückgabe: (abs_path, rel_path, year_label_guess)
    """
    if file_storage and getattr(file_storage, "filename", None):
        path, rel = _save_uploaded_file(
            file_storage, data.get("year_label") or "excel", (".xlsx",)
        )
        return path, rel, data.get("year_label") or "excel"

    rel = (data.get("excel_rel_path") or "").strip()
    if rel:
        path = _resolve_output_rel(rel)
        if not os.path.isfile(path) or not path.lower().endswith(".xlsx"):
            raise FileNotFoundError("Excel nicht gefunden oder ungültig.")
        # Jahr aus Pfad raten (erstes Segment wenn 4 Ziffern)
        parts = _normalize_rel_path(rel).split("/")
        year_guess = parts[0] if parts and re.fullmatch(r"\d{4}", parts[0]) else "export"
        return path, _normalize_rel_path(rel), year_guess

    year_label = re.sub(r"[^\w\-]", "", (data.get("year_label") or data.get("year") or ""))
    if year_label:
        path = _excel_path_for_year(year_label)
        if not os.path.exists(path):
            raise FileNotFoundError(
                f"Excel nicht gefunden: Rechnungsaufstellung_{year_label}.xlsx"
            )
        rel = f"{year_label}/Rechnungsaufstellung_{year_label}.xlsx"
        return path, rel, year_label

    raise ValueError("Kein Excel angegeben (Auswahl, Upload oder Jahr).")


def _serialize_match_result(result: dict) -> dict:
    """JSON-tauglich machen (date → iso)."""
    def _ser_match(m):
        out = dict(m)
        d = out.get("date")
        if hasattr(d, "isoformat"):
            out["date"] = d.isoformat()
        return out

    def _ser_inv(inv):
        out = dict(inv)
        d = out.get("re_dat")
        if hasattr(d, "isoformat"):
            out["re_dat"] = d.isoformat()
        return out

    return {
        "matches": [_ser_match(m) for m in result.get("matches", [])],
        "ambiguous": [_ser_inv(a) for a in result.get("ambiguous", [])],
        "unmatched": [_ser_inv(u) for u in result.get("unmatched", [])],
        "no_amount": [_ser_inv(n) for n in result.get("no_amount", [])],
        "stats": result.get("stats", {}),
    }


@app.route("/api/bank-csv/files")
def api_bank_csv_files():
    """
    GET /api/bank-csv/files?type=csv|xlsx
    Listet Dateien unter OUTPUT_DIR (Issue #28 / #29).
    """
    kind = (request.args.get("type") or "csv").lower().strip()
    if kind in ("csv", "kontoauszug"):
        exts = (".csv",)
    elif kind in ("xlsx", "excel", "xls"):
        exts = (".xlsx",)
    else:
        return jsonify({"error": "type muss csv oder xlsx sein."}), 400
    try:
        files = _list_output_files(exts, max_depth=3)
        return jsonify({"files": files, "type": "csv" if exts == (".csv",) else "xlsx"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/bank-csv/preview", methods=["POST"])
def api_bank_csv_preview():
    """CSV aus Server-Pfad oder Upload → Header, Preset, Sample."""
    if request.content_type and "multipart/form-data" in (request.content_type or ""):
        data = request.form.to_dict()
        f = request.files.get("file")
    else:
        data = request.get_json(force=True, silent=True) or {}
        f = None
    try:
        path, rel = _resolve_csv_path(data, f)
        preview = csv_preview(path)
        preview["csv_rel_path"] = rel
        preview["saved_path"] = os.path.basename(path)
        preview["upload_key"] = os.path.basename(path) if rel.startswith("_uploads/") else ""
        return jsonify(preview)
    except (ValueError, FileNotFoundError, PermissionError) as e:
        code = 403 if isinstance(e, PermissionError) else 400
        return jsonify({"error": str(e)}), code
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/bank-csv/match", methods=["POST"])
def api_bank_csv_match():
    """
    Matching (dry_run oder apply).

    csv_rel_path | file | upload_key
    excel_rel_path | excel_file | year_label (Fallback)
    """
    if request.content_type and "multipart/form-data" in (request.content_type or ""):
        data = request.form.to_dict()
        f_csv = request.files.get("file")
        f_excel = request.files.get("excel_file")
    else:
        data = request.get_json(force=True, silent=True) or {}
        f_csv = None
        f_excel = None

    dry_run = str(data.get("dry_run", "true")).lower() in ("1", "true", "yes")
    debits_only = str(data.get("debits_only", "true")).lower() in ("1", "true", "yes")
    only_relevant = str(data.get("only_relevant", "false")).lower() in ("1", "true", "yes")

    try:
        excel_path, excel_rel, year_label = _resolve_excel_path(data, f_excel)
        csv_path, csv_rel = _resolve_csv_path(data, f_csv)

        bank_rows = parse_bank_csv(
            csv_path, debits_only=debits_only, only_relevant=only_relevant
        )
        invoices = read_invoices_for_matching(excel_path)
        result = match_invoices_to_bank(invoices, bank_rows)

        excel_updated = 0
        if not dry_run:
            excel_updated = update_excel_with_bank_matches(excel_path, result)
            with job_lock:
                job_status["excel_path"] = excel_path

        payload = _serialize_match_result(result)
        payload["dry_run"] = dry_run
        payload["excel_path"] = excel_path
        payload["excel_rel_path"] = excel_rel
        payload["csv_rel_path"] = csv_rel
        payload["year_label"] = year_label
        payload["excel_updated"] = excel_updated
        payload["bank_rows_total"] = len(bank_rows)
        payload["bank_rows_selected"] = sum(1 for b in bank_rows if b.get("selected"))
        return jsonify(payload)
    except PermissionError as e:
        return jsonify({"error": str(e)}), 403
    except (ValueError, FileNotFoundError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/bank-csv/stb-export", methods=["POST"])
def api_bank_csv_stb_export():
    """
    STB-Export-Modus: Farben/Kommentare entfernen, Match-Status leeren + ausblenden,
    gefüllte Zahlungsdaten behalten (Issue #34).
    Body: excel_rel_path | excel_file | year_label
    """
    if request.content_type and "multipart/form-data" in (request.content_type or ""):
        data = request.form.to_dict()
        f_excel = request.files.get("excel_file")
    else:
        data = request.get_json(force=True, silent=True) or {}
        f_excel = None
    try:
        excel_path, excel_rel, year_label = _resolve_excel_path(data, f_excel)
        stats = prepare_stb_export(excel_path)
        return jsonify({
            "ok": True,
            "excel_path": excel_path,
            "excel_rel_path": excel_rel,
            "year_label": year_label,
            **stats,
        })
    except PermissionError as e:
        return jsonify({"error": str(e)}), 403
    except (ValueError, FileNotFoundError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
