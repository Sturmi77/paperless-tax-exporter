"""
Excel-Export im Format der Steuerberater-Vorlage.
Struktur: Tabelle1 mit 10 Spalten, Header in Zeile 4, SUMME in Zeile 1.
Spalten:
  A: Beleg-Nr.
  B: Re-Dat
  C: Zahlungsdatum  (leer – manuell zu befüllen)
  D: Bar / Konto / KK (leer – manuell zu befüllen)
  E: Absender (correspondent aus Paperless oder OCR-Vorschlag gelb)
  F: Beschreibung
  G: Kennzahl (document_type)
  H: Rechnungssumme (leer Stufe 1 / OCR-Vorschlag Stufe 2 gelb)
  I: Rechnungssumme inkl. Privatanteil (leer – manuell)
  J: Dateiname (Hyperlink zur PDF)
"""

import os
from datetime import datetime, date
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# Farben – komm|event CI
COLOR_HEADER_BG   = "2997AB"   # Teal (komm|event Primärfarbe)
COLOR_HEADER_FONT = "FFFFFF"   # Weiß
COLOR_SUM_BG      = "D6EEF2"   # Teal-hell (abgeleitet)
COLOR_SUM_FONT    = "1A6478"   # Teal-dunkel für SUMME-Beschriftung
COLOR_OCR_BG      = "FFFFC7"   # Gelb für OCR-Vorschlagswerte (funktional, bleibt)
COLOR_EMPTY_BG    = "F2F2F2"   # Hellgrau für manuell zu füllende Felder
COLOR_HYPERLINK   = "1E7D8F"   # Teal-dunkel für Hyperlinks

THIN   = Side(style="thin", color="C5D2D4")  # CI-Border
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLUMNS = [
    ("Beleg-Nr.",                           5.8,   "center"),
    ("Re-Dat",                              12.0,  "center"),
    ("Zahlungs-\ndatum",                    12.0,  "center"),
    ("Bar / Konto / KK",                    13.0,  "center"),
    ("Absender",                            28.0,  "left"),
    ("Beschreibung",                        38.0,  "left"),
    ("Kennzahl",                            22.0,  "left"),
    ("Rechnungssumme",                      20.0,  "right"),
    ("Rechnungs-\nsumme inkl. Privatanteil", 18.0, "right"),
    ("Dateiname / Beleg",                   38.0,  "left"),
]

# Optionale Spalte K (INCLUDE_TEXT_PATH=true)
COLUMN_TEXT_PATH = ("Pfad (kopierbar)", 52.0, "left")

# Issue #25: Bank-Match-Spalten (werden bei Bedarf angehängt)
COLUMN_BOOKING_TEXT = ("Buchungstext", 42.0, "left")
COLUMN_MATCH_STATUS = ("Match-Status", 16.0, "center")

COLOR_MATCH_OK      = "C6EFCE"  # Grün – gefunden
COLOR_MATCH_AMBIG   = "FFFFC7"  # Gelb – mehrdeutig (wie OCR)
COLOR_MATCH_MISS    = "FCE4D6"  # Orange – nicht gefunden

DATE_FORMAT   = "DD.MM.YYYY"
NUMBER_FORMAT = '#,##0.00 "€"'


def _header_cell(ws, row, col, value, align="center"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(bold=True, color=COLOR_HEADER_FONT, size=10)
    cell.fill      = PatternFill("solid", fgColor=COLOR_HEADER_BG)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border    = BORDER
    return cell


def _data_cell(ws, row, col, value=None, align="left", number_fmt=None,
               bold=False, bg=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(size=10, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border    = BORDER
    if number_fmt:
        cell.number_format = number_fmt
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    return cell


def _make_comment(text):
    try:
        from openpyxl.comments import Comment
        return Comment(text, "Paperless Exporter")
    except Exception:
        return None


def _build_unc_path(unc_base, year_label, filename, subfolder: str = ""):
    """
    Baut Windows-UNC-Pfad für Excel-Hyperlink.
    unc_base:   z.B. \\\\SynologyDS923\\downloads\\steuerberater
    year_label: z.B. 2024
    filename:   z.B. 0012_Telekom.pdf
    subfolder:  optionaler Unterordner (Allowlist-validiert, default="")
    Ergebnis (ohne subfolder): \\\\SynologyDS923\\downloads\\steuerberater\\2024\\Belege\\0012_Telekom.pdf
    Ergebnis (mit subfolder):  \\\\SynologyDS923\\downloads\\steuerberater\\2024\\Archiv\\Belege\\0012_Telekom.pdf
    """
    if not unc_base or not filename:
        return filename or ""
    # Backslashes normalisieren
    base = unc_base.rstrip("\\")
    if subfolder:
        return f"{base}\\{year_label}\\{subfolder}\\Belege\\{filename}"
    return f"{base}\\{year_label}\\Belege\\{filename}"


def _build_cell_formula(subfolder_path: str) -> str:
    r"""
    Baut eine portable CELL("filename")-Hyperlink-Formel.

    Die Formel berechnet den Hyperlink-Pfad dynamisch zur Laufzeit:
      LEFT(CELL("filename"), FIND("[", CELL("filename"))-1) -> Ordnerpfad der Excel-Datei
      & "Belege\\datei.pdf" -> relativer Unterordner + Dateiname

    Dadurch friert Excel den Pfad NICHT ein - er bleibt nach Ordner-Verschiebung gueltig.

    WICHTIG: Funktioniert nur wenn die Datei bereits gespeichert ist.
    In einer neuen, ungespeicherten Datei gibt CELL("filename") "" zurueck.

    subfolder_path: relativer Teilpfad nach dem Ordner der Excel-Datei
                    z.B. "Belege\\0012_Telekom.pdf"
    """
    # Backslashes in der Formel müssen als \\ (4 Backslashes) codiert werden,
    # damit Excel 2 Backslashes interpretiert → ein Backslash im Pfad
    escaped = subfolder_path.replace("\\", "\\\\")
    return (
        f'=HYPERLINK('
        f'LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&"{escaped}",'
        f'"{subfolder_path.split(chr(92))[-1]}")'
    )


def create_excel(documents, pdf_map, output_path, year_label,
                 unc_base=None, ocr_results=None, subfolder: str = "",
                 hyperlink_mode: str = "cell", include_text_path: bool = False,
                 progress_fn=None):
    """
    Erstellt die Excel-Datei im Steuerberater-Format (Stufe 1).

    documents:         Liste von Paperless-Dokumenten (API-Dicts)
    pdf_map:           {doc_id: filename_in_pdf_folder}
    output_path:       Zielpfad der .xlsx-Datei
    year_label:        z.B. "2024"
    unc_base:          Windows-UNC-Pfad Basis (optional)
    ocr_results:       {doc_id: {"absender": ..., "betrag": ...}} (optional, Stufe 2)
    subfolder:         optionaler Unterordner (Allowlist-validiert, default="")
    hyperlink_mode:    "cell" = CELL()-Formel (portabel); "unc" = absoluter UNC-Pfad (Issue #8)
    include_text_path: True = Spalte K mit kopierbarem UNC-Pfad (Issue #8)
    progress_fn:       optional callable(idx, total, title) – Fortschritt pro Zeile (Issue #19)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rechnungsaufstellung"

    # ── Zeile 1: SUMME-Zeile ──────────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    sum_label = ws.cell(row=1, column=1, value=f"SUMME {year_label}")
    sum_label.font      = Font(bold=True, size=11, color=COLOR_SUM_FONT)
    sum_label.fill      = PatternFill("solid", fgColor=COLOR_SUM_BG)
    sum_label.alignment = Alignment(horizontal="left", vertical="center")
    sum_label.border    = BORDER

    # ── Zeilen 2–3: leer ─────────────────────────────────────────────────
    ws.row_dimensions[2].height = 8
    ws.row_dimensions[3].height = 8

    # ── Zeile 4: Spaltenheader ───────────────────────────────────────────
    ws.row_dimensions[4].height = 36
    active_columns = list(COLUMNS)
    if include_text_path:
        active_columns.append(COLUMN_TEXT_PATH)
    for col_idx, (header, width, align) in enumerate(active_columns, start=1):
        _header_cell(ws, 4, col_idx, header, align)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Daten ab Zeile 5 ─────────────────────────────────────────────────
    data_start_row = 5
    sorted_docs = sorted(
        documents,
        key=lambda d: d.get("created", "1900-01-01") or "1900-01-01"
    )
    ocr = ocr_results or {}
    total_docs = len(sorted_docs)

    for i, doc in enumerate(sorted_docs):
        row = data_start_row + i
        ws.row_dimensions[row].height = 18
        doc_id = doc.get("id")
        if progress_fn:
            progress_fn(i + 1, total_docs, doc.get("title", f"Dokument {doc_id}"))

        # A: Beleg-Nr.
        beleg_nr = doc.get("archive_serial_number") or doc_id
        _data_cell(ws, row, 1, beleg_nr, align="center")

        # B: Re-Dat
        created_str = doc.get("created", "")
        dt = None
        if created_str:
            try:
                dt = datetime.strptime(created_str[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
        _data_cell(ws, row, 2, dt, align="center", number_fmt=DATE_FORMAT)

        # C: Zahlungsdatum – leer, manuell
        _data_cell(ws, row, 3, None, align="center",
                   number_fmt=DATE_FORMAT, bg=COLOR_EMPTY_BG)

        # D: Bar / Konto / KK – leer, manuell
        _data_cell(ws, row, 4, None, align="center", bg=COLOR_EMPTY_BG)

        # E: Absender
        ocr_data    = ocr.get(doc_id, {})
        correspondent = doc.get("correspondent_name") or None
        ocr_absender  = ocr_data.get("absender")

        if correspondent:
            # Aus Paperless – zuverlässig
            _data_cell(ws, row, 5, correspondent, align="left")
        elif ocr_absender:
            # OCR-Vorschlag – gelb
            cell_e = _data_cell(ws, row, 5, ocr_absender, align="left",
                                 bg=COLOR_OCR_BG)
            cell_e.comment = _make_comment("OCR-Vorschlag – bitte prüfen!")
        else:
            _data_cell(ws, row, 5, None, align="left", bg=COLOR_EMPTY_BG)

        # F: Beschreibung
        _data_cell(ws, row, 6, doc.get("title", ""), align="left")

        # G: Kennzahl (Dokumenttyp)
        doc_type_name = (doc.get("document_type_name")
                         or str(doc.get("document_type", "")) or "")
        _data_cell(ws, row, 7, doc_type_name, align="left")

        # H: Rechnungssumme
        ocr_betrag = ocr_data.get("betrag")
        if ocr_betrag is not None:
            cell_h = _data_cell(ws, row, 8, ocr_betrag, align="right",
                                 number_fmt=NUMBER_FORMAT, bg=COLOR_OCR_BG)
            cell_h.comment = _make_comment("OCR-Vorschlag – bitte prüfen!")
        else:
            _data_cell(ws, row, 8, None, align="right",
                       number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # I: Rechnungssumme inkl. Privatanteil – leer, manuell
        _data_cell(ws, row, 9, None, align="right",
                   number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # J: Dateiname / Hyperlink (Issue #8: CELL-Formel oder UNC)
        filename = pdf_map.get(doc_id, "")
        if filename:
            if hyperlink_mode == "cell":
                # CELL("filename")-Formel: portabel, funktioniert nach Ordner-Verschiebung
                if subfolder:
                    rel_path = f"{subfolder}\\Belege\\{filename}"
                else:
                    rel_path = f"Belege\\{filename}"
                formula = _build_cell_formula(rel_path)
            elif unc_base:
                # Absoluter UNC-Pfad (Fallback / HYPERLINK_MODE=unc)
                unc_path = _build_unc_path(unc_base, year_label, filename, subfolder)
                formula  = f'=HYPERLINK("{unc_path}","{filename}")'
            else:
                formula = filename

            cell_j = ws.cell(row=row, column=10, value=formula)
            cell_j.font      = Font(size=10, color=COLOR_HYPERLINK, underline="single")
            cell_j.alignment = Alignment(horizontal="left", vertical="center")
            cell_j.border    = BORDER

            # Spalte K: kopierbarer UNC-Pfad (INCLUDE_TEXT_PATH=true)
            if include_text_path and unc_base:
                unc_path = _build_unc_path(unc_base, year_label, filename, subfolder)
                cell_k = ws.cell(row=row, column=11, value=unc_path)
                cell_k.font      = Font(size=9, color="666666")
                cell_k.alignment = Alignment(horizontal="left", vertical="center")
                cell_k.border    = BORDER
        else:
            _data_cell(ws, row, 10, filename, align="left")

    last_data_row = data_start_row + len(sorted_docs) - 1

    # ── Excel-Tabelle (Tabelle1) ──────────────────────────────────────────
    num_cols  = len(COLUMNS) + (1 if include_text_path else 0)
    table_ref = f"A4:{get_column_letter(num_cols)}{last_data_row}"
    table = Table(displayName="Tabelle1", ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True,  showColumnStripes=False,
    )
    ws.add_table(table)

    # Summenformel (Spalte H = Rechnungssumme)
    ws["H1"] = f"=SUM(H{data_start_row}:H{last_data_row})"
    ws["H1"].number_format = NUMBER_FORMAT
    ws["H1"].font          = Font(bold=True, size=11, color=COLOR_SUM_FONT)
    ws["H1"].fill          = PatternFill("solid", fgColor=COLOR_SUM_BG)
    ws["H1"].alignment     = Alignment(horizontal="right", vertical="center")
    ws["H1"].border        = BORDER

    ws.freeze_panes = "A5"
    ws.print_area   = f"A1:{get_column_letter(len(COLUMNS))}{last_data_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True

    wb.save(output_path)
    return output_path


def update_excel_with_ocr(excel_path, ocr_results, unc_base, year_label,
                          subfolder: str = ""):
    """
    Stufe 2: Öffnet bestehendes Excel und trägt OCR-Ergebnisse ein.
    Überschreibt nur leere oder bereits gelbe Felder (schützt manuelle Einträge).

    ocr_results: {doc_id: {"absender": str|None, "betrag": float|None}}
    subfolder:   optionaler Unterordner unterhalb des Jahres-Ordners (default="")
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel nicht gefunden: {excel_path}")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Rechnungsaufstellung"]

    YELLOW_FILL = PatternFill("solid", fgColor=COLOR_OCR_BG)
    EMPTY_FILL  = PatternFill("solid", fgColor=COLOR_EMPTY_BG)
    NONE_FILL   = PatternFill(fill_type=None)

    updated = 0

    # Daten starten ab Zeile 5 – Spalte A = Beleg-Nr. (nutzen wir als Key)
    # Wir brauchen eine Zuordnung Beleg-Nr./ID → Zeile
    # Da doc_id in Spalte A steht (ASN oder ID), iterieren wir

    # Baue Mapping: Wert in Spalte A → Zeile
    row_map = {}
    for row in ws.iter_rows(min_row=5):
        cell_a = row[0]
        if cell_a.value is not None:
            row_map[str(cell_a.value)] = cell_a.row

    for doc_id, data in ocr_results.items():
        # doc_id könnte ASN oder numerische ID sein – probiere beide
        row_num = row_map.get(str(doc_id))
        if row_num is None:
            continue

        absender = data.get("absender")
        betrag   = data.get("betrag")

        # Spalte E (Absender) – nur überschreiben wenn leer oder gelb
        cell_e = ws.cell(row=row_num, column=5)
        e_is_empty = cell_e.value is None or cell_e.value == ""
        e_is_ocr   = (cell_e.fill and cell_e.fill.fgColor and
                      cell_e.fill.fgColor.rgb == COLOR_OCR_BG)
        if absender and (e_is_empty or e_is_ocr):
            cell_e.value      = absender
            cell_e.fill       = YELLOW_FILL
            cell_e.font       = Font(size=10)
            cell_e.alignment  = Alignment(horizontal="left", vertical="center")
            cell_e.border     = BORDER
            cell_e.comment    = _make_comment("OCR-Vorschlag – bitte prüfen!")
            updated += 1

        # Spalte H (Rechnungssumme) – nur überschreiben wenn leer oder gelb
        cell_h = ws.cell(row=row_num, column=8)
        h_is_empty = cell_h.value is None or cell_h.value == ""
        h_is_ocr   = (cell_h.fill and cell_h.fill.fgColor and
                      cell_h.fill.fgColor.rgb == COLOR_OCR_BG)
        if betrag is not None and (h_is_empty or h_is_ocr):
            cell_h.value         = betrag
            cell_h.fill          = YELLOW_FILL
            cell_h.font          = Font(size=10)
            cell_h.alignment     = Alignment(horizontal="right", vertical="center")
            cell_h.number_format = NUMBER_FORMAT
            cell_h.border        = BORDER
            cell_h.comment       = _make_comment("OCR-Vorschlag – bitte prüfen!")
            updated += 1

        # Spalte J (Hyperlink) – nur wenn noch kein Hyperlink vorhanden
        cell_j = ws.cell(row=row_num, column=10)
        filename = str(cell_j.value or "")
        if filename and unc_base and not str(filename).startswith("=HYPERLINK"):
            unc_path = _build_unc_path(unc_base, year_label, filename, subfolder)
            cell_j.value      = f'=HYPERLINK("{unc_path}","{filename}")'
            cell_j.font       = Font(size=10, color=COLOR_HYPERLINK, underline="single")
            cell_j.alignment  = Alignment(horizontal="left", vertical="center")
            cell_j.border     = BORDER

    wb.save(excel_path)
    return updated


def get_existing_doc_ids(excel_path):
    """
    Liest alle Beleg-Nummern (Spalte A ab Zeile 5) aus einem bestehenden Excel.
    Gibt ein Set von Strings zurück (ASN oder numerische ID).
    """
    if not os.path.exists(excel_path):
        return set()
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb["Rechnungsaufstellung"]
    ids = set()
    for row in ws.iter_rows(min_row=5, max_col=1, values_only=True):
        val = row[0]
        if val is not None:
            ids.add(str(val))
    wb.close()
    return ids


def append_to_excel(new_documents, pdf_map, excel_path, year_label,
                    unc_base=None, subfolder: str = "",
                    hyperlink_mode: str = "cell", include_text_path: bool = False):
    """
    Hängt neue Dokumente an ein bestehendes Excel an.
    Bestehende Zeilen werden nicht verändert.
    Gibt die Anzahl neu angehängter Zeilen zurück.

    new_documents:     Liste von Paperless-Dokumenten die noch NICHT im Excel sind
    pdf_map:           {doc_id: filename}
    subfolder:         optionaler Unterordner (default="")
    hyperlink_mode:    "cell" = CELL()-Formel; "unc" = absoluter UNC-Pfad (Issue #8)
    include_text_path: Spalte K mit kopierbarem UNC-Pfad (Issue #8)
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel nicht gefunden: {excel_path}")
    if not new_documents:
        return 0

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Rechnungsaufstellung"]

    # Letzte belegte Zeile in Spalte A finden
    last_row = 4  # Mindestens Header-Zeile
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value is not None:
            last_row = row[0].row

    insert_start = last_row + 1

    sorted_new = sorted(
        new_documents,
        key=lambda d: d.get("created", "1900-01-01") or "1900-01-01"
    )

    for i, doc in enumerate(sorted_new):
        row = insert_start + i
        ws.row_dimensions[row].height = 18
        doc_id = doc.get("id")

        # A: Beleg-Nr.
        beleg_nr = doc.get("archive_serial_number") or doc_id
        _data_cell(ws, row, 1, beleg_nr, align="center")

        # B: Re-Dat
        created_str = doc.get("created", "")
        dt = None
        if created_str:
            try:
                from datetime import datetime as _dt
                dt = _dt.strptime(created_str[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
        _data_cell(ws, row, 2, dt, align="center", number_fmt=DATE_FORMAT)

        # C: Zahlungsdatum
        _data_cell(ws, row, 3, None, align="center",
                   number_fmt=DATE_FORMAT, bg=COLOR_EMPTY_BG)
        # D: Bar / Konto / KK
        _data_cell(ws, row, 4, None, align="center", bg=COLOR_EMPTY_BG)

        # E: Absender
        correspondent = doc.get("correspondent_name") or None
        if correspondent:
            _data_cell(ws, row, 5, correspondent, align="left")
        else:
            _data_cell(ws, row, 5, None, align="left", bg=COLOR_EMPTY_BG)

        # F: Beschreibung
        _data_cell(ws, row, 6, doc.get("title", ""), align="left")

        # G: Kennzahl
        doc_type_name = (doc.get("document_type_name")
                         or str(doc.get("document_type", "")) or "")
        _data_cell(ws, row, 7, doc_type_name, align="left")

        # H: Rechnungssumme (leer – kann später per Stufe 2 befüllt werden)
        _data_cell(ws, row, 8, None, align="right",
                   number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # I: Rechnungssumme inkl. Privatanteil
        _data_cell(ws, row, 9, None, align="right",
                   number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # J: Dateiname / Hyperlink (Issue #8)
        filename = pdf_map.get(doc_id, "")
        if filename:
            if hyperlink_mode == "cell":
                if subfolder:
                    rel_path = f"{subfolder}\\Belege\\{filename}"
                else:
                    rel_path = f"Belege\\{filename}"
                formula = _build_cell_formula(rel_path)
            elif unc_base:
                unc_path = _build_unc_path(unc_base, year_label, filename, subfolder)
                formula  = f'=HYPERLINK("{unc_path}","{filename}")'
            else:
                formula = filename

            cell_j = ws.cell(row=row, column=10, value=formula)
            cell_j.font      = Font(size=10, color=COLOR_HYPERLINK, underline="single")
            cell_j.alignment = Alignment(horizontal="left", vertical="center")
            cell_j.border    = BORDER

            if include_text_path and unc_base:
                unc_path = _build_unc_path(unc_base, year_label, filename, subfolder)
                cell_k = ws.cell(row=row, column=11, value=unc_path)
                cell_k.font      = Font(size=9, color="666666")
                cell_k.alignment = Alignment(horizontal="left", vertical="center")
                cell_k.border    = BORDER
        else:
            _data_cell(ws, row, 10, filename, align="left")

    # SUMME-Formel auf neue letzte Zeile ausdehnen
    new_last_row = insert_start + len(sorted_new) - 1
    ws["H1"] = f"=SUM(H5:H{new_last_row})"
    ws["H1"].number_format = NUMBER_FORMAT
    ws["H1"].font          = Font(bold=True, size=11, color=COLOR_SUM_FONT)
    ws["H1"].fill          = PatternFill("solid", fgColor=COLOR_SUM_BG)
    ws["H1"].alignment     = Alignment(horizontal="right", vertical="center")
    ws["H1"].border        = BORDER

    # Tabellen-Referenz ausweiten
    for tbl in ws.tables.values():
        if tbl.displayName == "Tabelle1":
            tbl.ref = f"A4:{get_column_letter(len(COLUMNS))}{new_last_row}"
            break

    wb.save(excel_path)
    return len(sorted_new)


# ── Issue #25: Kontoauszug-Matching ───────────────────────────────────

def _header_map(ws) -> dict:
    """Header-Zeile 4 → {normalisierter Name: Spaltenindex 1-basiert}."""
    result = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=4, column=col).value
        if val is None:
            continue
        key = str(val).replace("\n", " ").strip().lower()
        result[key] = col
    return result


def _ensure_bank_match_columns(ws) -> tuple:
    """
    Stellt sicher, dass Spalten Buchungstext und Match-Status existieren.
    Rückgabe: (col_booking_text, col_match_status)
    """
    headers = _header_map(ws)
    col_text = None
    col_status = None
    for name, col in headers.items():
        if "buchungstext" in name:
            col_text = col
        if "match-status" in name or name == "match status":
            col_status = col

    next_col = ws.max_column + 1
    # Leere trailing Spalten vermeiden: max_column anhand Header neu bestimmen
    last_header = 0
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=4, column=col).value is not None:
            last_header = col
    next_col = last_header + 1

    if col_text is None:
        col_text = next_col
        _header_cell(ws, 4, col_text, COLUMN_BOOKING_TEXT[0], COLUMN_BOOKING_TEXT[2])
        ws.column_dimensions[get_column_letter(col_text)].width = COLUMN_BOOKING_TEXT[1]
        next_col += 1
    if col_status is None:
        col_status = next_col
        _header_cell(ws, 4, col_status, COLUMN_MATCH_STATUS[0], COLUMN_MATCH_STATUS[2])
        ws.column_dimensions[get_column_letter(col_status)].width = COLUMN_MATCH_STATUS[1]

    return col_text, col_status


def _parse_excel_date(val):
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def read_invoices_for_matching(excel_path) -> list:
    """
    Liest Rechnungszeilen mit **leerem Zahlungsdatum (C)** fürs Matching.
    Rückgabe: [{row, beleg_nr, re_dat, absender, beschreibung, betrag}, ...]
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel nicht gefunden: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["Rechnungsaufstellung"]
    invoices = []
    for row in range(5, ws.max_row + 1):
        beleg = ws.cell(row=row, column=1).value
        if beleg is None or beleg == "":
            continue
        zahlung = ws.cell(row=row, column=3).value
        if zahlung is not None and zahlung != "":
            continue  # Entscheidung #2: nur leeres C
        re_dat = _parse_excel_date(ws.cell(row=row, column=2).value)
        absender = ws.cell(row=row, column=5).value
        beschreibung = ws.cell(row=row, column=6).value
        betrag = ws.cell(row=row, column=8).value
        try:
            betrag_f = float(betrag) if betrag is not None and betrag != "" else None
        except (TypeError, ValueError):
            betrag_f = None
        invoices.append({
            "row": row,
            "beleg_nr": beleg,
            "re_dat": re_dat,
            "absender": str(absender).strip() if absender else "",
            "beschreibung": str(beschreibung).strip() if beschreibung else "",
            "betrag": betrag_f,
        })
    wb.close()
    return invoices


def update_excel_with_bank_matches(excel_path, match_result: dict) -> int:
    """
    Schreibt Matching-Ergebnis ins Rechnungs-Excel (Issue #25).
    - gefunden: C=Datum, Buchungstext, Status grün
    - mehrdeutig: Status gelb + Kommentar, C unberührt
    - nicht gefunden: Status orange
    Manuell befüllte C-Zellen werden nicht überschrieben.
    """
    from matching_csv import STATUS_FOUND, STATUS_AMBIGUOUS, STATUS_NOT_FOUND

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel nicht gefunden: {excel_path}")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Rechnungsaufstellung"]
    col_text, col_status = _ensure_bank_match_columns(ws)
    updated = 0

    def _set_status(row, status, fill_color, comment=None):
        cell = ws.cell(row=row, column=col_status, value=status)
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.font = Font(size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        if comment:
            cell.comment = _make_comment(comment)

    for m in match_result.get("matches", []):
        row = m["invoice_row"]
        cell_c = ws.cell(row=row, column=3)
        if cell_c.value is None or cell_c.value == "":
            cell_c.value = m["date"]
            cell_c.number_format = DATE_FORMAT
            cell_c.fill = PatternFill("solid", fgColor=COLOR_MATCH_OK)
            cell_c.font = Font(size=10)
            cell_c.alignment = Alignment(horizontal="center", vertical="center")
            cell_c.border = BORDER
            cell_c.comment = _make_comment("Aus Kontoauszug zugeordnet – bitte prüfen!")
            updated += 1
        cell_t = ws.cell(row=row, column=col_text, value=m.get("text") or "")
        cell_t.fill = PatternFill("solid", fgColor=COLOR_MATCH_OK)
        cell_t.font = Font(size=10)
        cell_t.alignment = Alignment(horizontal="left", vertical="center")
        cell_t.border = BORDER
        _set_status(row, STATUS_FOUND, COLOR_MATCH_OK)
        updated += 1

    for a in match_result.get("ambiguous", []):
        row = a["row"]
        cands = a.get("candidates") or []
        lines = []
        for c in cands[:3]:
            lines.append(
                f"#{c.get('bank_row_id')} {c.get('date')} {c.get('amount')} "
                f"{(c.get('partner') or '')[:40]}"
            )
        comment = "Mehrdeutig:\n" + "\n".join(lines) if lines else "Mehrere Treffer"
        _set_status(row, STATUS_AMBIGUOUS, COLOR_MATCH_AMBIG, comment)
        updated += 1

    for u in match_result.get("unmatched", []):
        row = u["row"]
        _set_status(row, STATUS_NOT_FOUND, COLOR_MATCH_MISS)
        updated += 1

    # Tabellenbereich erweitern falls Tabelle1 existiert
    last_row = ws.max_row
    last_col = max(col_text, col_status)
    for tbl in ws.tables.values():
        if tbl.displayName == "Tabelle1":
            tbl.ref = f"A4:{get_column_letter(last_col)}{last_row}"
            break

    wb.save(excel_path)
    return updated


def write_filtered_bank_xlsx(bank_rows: list, match_result: dict, output_path: str) -> str:
    """
    Schreibt Kontoauszug_gefiltert_*.xlsx mit matched/ambiguous Zeilen (C3).
    Farben: grün = zugeordnet, gelb = mehrdeutig.
    """
    from matching_csv import STATUS_FOUND, STATUS_AMBIGUOUS

    used = set(match_result.get("used_bank_ids") or [])
    amb_ids = set()
    for a in match_result.get("ambiguous", []):
        for c in a.get("candidates") or []:
            amb_ids.add(c.get("bank_row_id"))

    match_by_bank = {m["bank_row_id"]: m for m in match_result.get("matches", [])}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Zugeordnet"
    headers = ["CSV-Zeile", "Datum", "Betrag", "Partner", "Buchungstext", "Status", "Beleg-Nr."]
    for i, h in enumerate(headers, start=1):
        _header_cell(ws, 1, i, h, "center")
        ws.column_dimensions[get_column_letter(i)].width = [10, 12, 12, 28, 48, 14, 12][i - 1]

    out_row = 2
    for b in bank_rows:
        rid = b["row_id"]
        if rid in used:
            status = STATUS_FOUND
            fill = COLOR_MATCH_OK
            beleg = match_by_bank.get(rid, {}).get("beleg_nr", "")
        elif rid in amb_ids:
            status = STATUS_AMBIGUOUS
            fill = COLOR_MATCH_AMBIG
            beleg = ""
        else:
            continue

        values = [
            rid,
            b.get("date"),
            b.get("amount"),
            b.get("partner") or "",
            (b.get("text") or "")[:200],
            status,
            beleg,
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=out_row, column=col, value=val)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = Font(size=10)
            cell.border = BORDER
            if col == 2 and val is not None:
                cell.number_format = DATE_FORMAT
            if col == 3 and val is not None:
                cell.number_format = NUMBER_FORMAT
        out_row += 1

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    return output_path


# Need date import at top if not present — check excel_export imports
