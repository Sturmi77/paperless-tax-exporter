"""
Kontoauszug-CSV parsen und normalisieren (Issue #25).
Preset „AT Giro“ für österreichische Banken-Exporte (Semikolon, deutsches Betragsformat).
"""

from __future__ import annotations

import csv
import io
import os
import re
from datetime import datetime, date
from typing import Optional


# Spalten-Mapping Preset „AT Giro“ (Referenz: umsaetze-girokonto_AT….csv)
AT_GIRO_PRESET = {
    "date": "Buchungsdatum",
    "date_fallback": "Valutadatum",
    "amount": "Betrag",
    "partner": "Name des Partners",
    "booking_type": "Buchungstext",
    "text_primary": "Verwendungszweck",
    "text_fallback": "Umsatztext",
    "relevant": "Relevant",  # optional Vorfilter-Spalte
}


def parse_de_amount(raw) -> Optional[float]:
    """Parst deutsches Betragsformat: '-1.400,00' / '4.952,37' / '-16,05'."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    s = s.replace(" ", "").replace("€", "").replace("EUR", "")
    # Tausenderpunkt entfernen, Komma → Punkt
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_iso_date(raw) -> Optional[date]:
    """Parst YYYY-MM-DD oder DD.MM.YYYY."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s[:10] if fmt.startswith("%Y") else s, fmt).date()
        except ValueError:
            continue
    # Buchungszeit mit Suffix: 2025-12-31-21.23.51…
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", s)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass
    return None


def detect_delimiter(sample: str) -> str:
    """Erkennt ; oder , als Delimiter anhand der Header-Zeile."""
    first = sample.splitlines()[0] if sample else ""
    if first.count(";") >= first.count(","):
        return ";"
    return ","


def _read_text(path: str) -> str:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _cell(row: dict, *keys: str) -> str:
    for k in keys:
        if not k:
            continue
        # exakter Key
        if k in row and row[k] is not None and str(row[k]).strip():
            return str(row[k]).strip()
        # case-insensitive
        for rk, rv in row.items():
            if rk and rk.strip().lower() == k.lower() and rv is not None and str(rv).strip():
                return str(rv).strip()
    return ""


def _is_at_giro_header(fieldnames: list) -> bool:
    names = {f.strip().lower() for f in (fieldnames or []) if f}
    return "buchungsdatum" in names and "betrag" in names


def parse_bank_csv(
    path: str,
    *,
    debits_only: bool = True,
    only_relevant: bool = False,
) -> list[dict]:
    """
    Liest Kontoauszug-CSV und gibt normalisierte Buchungen zurück.

    Jeder Eintrag:
      {
        row_id, date, amount, partner, booking_type, text,
        selected, raw
      }
    amount: vorzeichenbehaftet (negativ = Abbuchung)
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"CSV nicht gefunden: {path}")

    text = _read_text(path)
    delim = detect_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    if not reader.fieldnames:
        raise ValueError("CSV hat keine Header-Zeile.")

    use_preset = _is_at_giro_header(list(reader.fieldnames))
    preset = AT_GIRO_PRESET if use_preset else None

    rows = []
    for idx, raw in enumerate(reader, start=2):  # Zeile 1 = Header
        if preset:
            d = parse_iso_date(_cell(raw, preset["date"], preset["date_fallback"]))
            amount = parse_de_amount(_cell(raw, preset["amount"]))
            partner = _cell(raw, preset["partner"])
            booking_type = _cell(raw, preset["booking_type"])
            text_val = _cell(raw, preset["text_primary"]) or _cell(raw, preset["text_fallback"])
            relevant_raw = _cell(raw, preset["relevant"])
        else:
            # Generischer Fallback: erste sinnvolle Spalten heuristisch
            d = parse_iso_date(
                _cell(raw, "Buchungsdatum", "Valutadatum", "Datum", "Date", "Booking date")
            )
            amount = parse_de_amount(
                _cell(raw, "Betrag", "Amount", "Umsatz", "Wert")
            )
            partner = _cell(raw, "Name des Partners", "Partner", "Empfänger", "Auftraggeber")
            booking_type = _cell(raw, "Buchungstext", "Typ", "Art")
            text_val = (
                _cell(raw, "Verwendungszweck", "Umsatztext", "Text", "Beschreibung")
            )
            relevant_raw = _cell(raw, "Relevant", "Match", "X")

        if d is None and amount is None:
            continue

        selected = True
        if only_relevant:
            selected = relevant_raw.upper() in ("X", "1", "JA", "YES", "TRUE")
        if debits_only and amount is not None and amount >= 0:
            selected = False

        rows.append({
            "row_id": idx,
            "date": d,
            "amount": amount,
            "partner": partner,
            "booking_type": booking_type,
            "text": text_val or booking_type,
            "selected": selected,
            "relevant_mark": relevant_raw,
            "preset": "at_giro" if use_preset else "generic",
        })

    return rows


def csv_preview(path: str, max_rows: int = 5) -> dict:
    """Preview für UI: Header, Delimiter, Preset, Sample-Zeilen."""
    text = _read_text(path)
    delim = detect_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    fieldnames = list(reader.fieldnames or [])
    sample = []
    for i, row in enumerate(reader):
        if i >= max_rows:
            break
        sample.append({k: (v if v is not None else "") for k, v in row.items()})
    return {
        "delimiter": delim,
        "fieldnames": fieldnames,
        "preset": "at_giro" if _is_at_giro_header(fieldnames) else "generic",
        "sample": sample,
        "row_count_estimate": text.count("\n"),
    }
