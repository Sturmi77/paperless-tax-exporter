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
from datetime import datetime
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# Farben
COLOR_HEADER_BG   = "1F497D"   # Dunkelblau für Header
COLOR_HEADER_FONT = "FFFFFF"   # Weiß
COLOR_SUM_BG      = "DCE6F1"   # Hellblau für Summenzeile
COLOR_OCR_BG      = "FFFFC7"   # Gelb für OCR-Vorschlagswerte
COLOR_EMPTY_BG    = "F2F2F2"   # Hellgrau für manuell zu füllende Felder

THIN   = Side(style="thin", color="BFBFBF")
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


def _build_unc_path(unc_base, year_label, filename):
    """
    Baut Windows-UNC-Pfad für Excel-Hyperlink.
    unc_base:   z.B. \\\\SynologyDS923\\downloads\\steuerberater
    year_label: z.B. 2024
    filename:   z.B. 0012_Telekom.pdf
    Ergebnis:   \\\\SynologyDS923\\downloads\\steuerberater\\2024\\Belege\\0012_Telekom.pdf
    """
    if not unc_base or not filename:
        return filename or ""
    # Backslashes normalisieren
    base = unc_base.rstrip("\\")
    return f"{base}\\{year_label}\\Belege\\{filename}"


def create_excel(documents, pdf_map, output_path, year_label,
                 unc_base=None, ocr_results=None):
    """
    Erstellt die Excel-Datei im Steuerberater-Format (Stufe 1).

    documents:   Liste von Paperless-Dokumenten (API-Dicts)
    pdf_map:     {doc_id: filename_in_pdf_folder}
    output_path: Zielpfad der .xlsx-Datei
    year_label:  z.B. "2024"
    unc_base:    Windows-UNC-Pfad Basis (optional)
    ocr_results: {doc_id: {"absender": ..., "betrag": ...}} (optional, Stufe 2)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rechnungsaufstellung"

    # ── Zeile 1: SUMME-Zeile ──────────────────────────────────────────────
    ws.row_dimensions[1].height = 20
    sum_label = ws.cell(row=1, column=1, value=f"SUMME {year_label}")
    sum_label.font      = Font(bold=True, size=11)
    sum_label.alignment = Alignment(horizontal="left", vertical="center")

    # ── Zeilen 2–3: leer ─────────────────────────────────────────────────
    ws.row_dimensions[2].height = 8
    ws.row_dimensions[3].height = 8

    # ── Zeile 4: Spaltenheader ───────────────────────────────────────────
    ws.row_dimensions[4].height = 36
    for col_idx, (header, width, align) in enumerate(COLUMNS, start=1):
        _header_cell(ws, 4, col_idx, header, align)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Daten ab Zeile 5 ─────────────────────────────────────────────────
    data_start_row = 5
    sorted_docs = sorted(
        documents,
        key=lambda d: d.get("created", "1900-01-01") or "1900-01-01"
    )
    ocr = ocr_results or {}

    for i, doc in enumerate(sorted_docs):
        row = data_start_row + i
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

        # J: Dateiname / Hyperlink
        filename = pdf_map.get(doc_id, "")
        if filename and unc_base:
            unc_path = _build_unc_path(unc_base, year_label, filename)
            # Excel HYPERLINK Formel
            cell_j = ws.cell(
                row=row, column=10,
                value=f'=HYPERLINK("{unc_path}","{filename}")'
            )
            cell_j.font      = Font(size=10, color="0563C1", underline="single")
            cell_j.alignment = Alignment(horizontal="left", vertical="center")
            cell_j.border    = BORDER
        else:
            _data_cell(ws, row, 10, filename, align="left")

    last_data_row = data_start_row + len(sorted_docs) - 1

    # ── Excel-Tabelle (Tabelle1) ──────────────────────────────────────────
    table_ref = f"A4:{get_column_letter(len(COLUMNS))}{last_data_row}"
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
    ws["H1"].font          = Font(bold=True, size=11)
    ws["H1"].alignment     = Alignment(horizontal="right", vertical="center")

    ws.freeze_panes = "A5"
    ws.print_area   = f"A1:{get_column_letter(len(COLUMNS))}{last_data_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True

    wb.save(output_path)
    return output_path


def update_excel_with_ocr(excel_path, ocr_results, unc_base, year_label):
    """
    Stufe 2: Öffnet bestehendes Excel und trägt OCR-Ergebnisse ein.
    Überschreibt nur leere oder bereits gelbe Felder (schützt manuelle Einträge).

    ocr_results: {doc_id: {"absender": str|None, "betrag": float|None}}
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
            unc_path = _build_unc_path(unc_base, year_label, filename)
            cell_j.value      = f'=HYPERLINK("{unc_path}","{filename}")'
            cell_j.font       = Font(size=10, color="0563C1", underline="single")
            cell_j.alignment  = Alignment(horizontal="left", vertical="center")
            cell_j.border     = BORDER

    wb.save(excel_path)
    return updated
