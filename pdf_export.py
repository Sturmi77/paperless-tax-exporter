"""
PDF-Download aus Paperless NGX.
Lädt jedes Dokument als PDF herunter und speichert es im Zielordner.
Dateiname: {ASN:04d}_{bereinigter_titel}.pdf  oder  {id}_{titel}.pdf
"""

import os
import re
import requests


def _sanitize_filename(name: str, max_len: int = 80) -> str:
    """Bereinigt einen String für Dateinamen (Windows + Linux sicher)."""
    # Ungültige Zeichen ersetzen
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    # Mehrfache Spaces/Underscores zusammenfassen
    name = re.sub(r"[\s_]+", "_", name).strip("_")
    # Länge begrenzen
    return name[:max_len]


def _make_pdf_filename(doc: dict) -> str:
    asn = doc.get("archive_serial_number")
    title = doc.get("title", f"dokument_{doc.get('id', '0')}")
    title_clean = _sanitize_filename(title)

    if asn:
        return f"{int(asn):04d}_{title_clean}.pdf"
    else:
        doc_id = doc.get("id", 0)
        return f"{int(doc_id):05d}_{title_clean}.pdf"


def download_pdfs(
    documents: list,
    pdf_folder: str,
    paperless_url: str,
    token: str,
    log_fn=None,
) -> dict:
    """
    Lädt alle PDFs herunter.

    Rückgabe: {doc_id: filename}  (nur Dateiname, nicht vollständiger Pfad)
    """
    headers = {"Authorization": f"Token {token}"}
    pdf_map = {}
    total = len(documents)

    for idx, doc in enumerate(documents, start=1):
        doc_id = doc.get("id")
        filename = _make_pdf_filename(doc)
        target_path = os.path.join(pdf_folder, filename)

        # Bereits vorhanden? Überspringen.
        if os.path.exists(target_path):
            if log_fn:
                log_fn(f"  [{idx}/{total}] Übersprungen (bereits vorhanden): {filename}")
            pdf_map[doc_id] = filename
            continue

        url = f"{paperless_url.rstrip('/')}/api/documents/{doc_id}/download/"
        try:
            resp = requests.get(url, headers=headers, timeout=60, stream=True)
            resp.raise_for_status()
            with open(target_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            pdf_map[doc_id] = filename
            if log_fn:
                log_fn(f"  [{idx}/{total}] ✓ {filename}")
        except requests.HTTPError as e:
            if log_fn:
                log_fn(f"  [{idx}/{total}] ✗ Fehler bei Dokument {doc_id}: {e}")
        except Exception as e:
            if log_fn:
                log_fn(f"  [{idx}/{total}] ✗ Unbekannter Fehler bei {doc_id}: {e}")

    return pdf_map
