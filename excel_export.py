"""
Excel-Export im Format der Steuerberater-Vorlage.
Struktur: Tabelle1 mit 9 Spalten, Header in Zeile 4, SUMME in Zeile 1.
Spalten:
  A: Beleg-Nr.
  B: Re-Dat
  C: Zahlungsdatum  (leer – manuell zu befüllen)
  D: Bar / Konto / KK (leer – manuell zu befüllen)
  E: Beschreibung
  F: Kennzahl (document_type)
  G: Rechnungssumme (leer Stufe 1 / OCR-Vorschlag Stufe 2)
  H: Rechnungssumme inkl. Privatanteil (leer – manuell)
  I: Dateiname (Link zur PDF)
"""

import os
from datetime import datetime
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY


# Farben
COLOR_HEADER_BG   = "1F497D"   # Dunkelblau für Header
COLOR_HEADER_FONT = "FFFFFF"   # Weiß
COLOR_SUM_BG      = "DCE6F1"   # Hellblau für Summenzeile
COLOR_OCR_BG      = "FFFFC7"   # Gelb für OCR-Vorschlagswerte (Stufe 2)
COLOR_EMPTY_BG    = "F2F2F2"   # Hellgrau für manuell zu füllende Felder

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLUMNS = [
    ("Beleg-Nr.",                        5.8,   "center"),
    ("Re-Dat",                           12.0,  "center"),
    ("Zahlungs-\ndatum",                 12.0,  "center"),
    ("Bar / Konto / KK",                 13.0,  "center"),
    ("Beschreibung",                     38.0,  "left"),
    ("Kennzahl",                         28.0,  "left"),
    ("Rechnungssumme",                   20.0,  "right"),
    ("Rechnungs-\nsumme inkl. Privatanteil", 18.0, "right"),
    ("Dateiname / Beleg",                30.0,  "left"),
]

DATE_FORMAT = "DD.MM.YYYY"
NUMBER_FORMAT = '#,##0.00 "€"'


def _header_cell(ws, row, col, value, align="center"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=True, color=COLOR_HEADER_FONT, size=10)
    cell.fill = PatternFill("solid", fgColor=COLOR_HEADER_BG)
    cell.alignment = Alignment(
        horizontal=align, vertical="center", wrap_text=True
    )
    cell.border = BORDER
    return cell


def _data_cell(ws, row, col, value=None, align="left", number_fmt=None,
               bold=False, bg=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(size=10, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = BORDER
    if number_fmt:
        cell.number_format = number_fmt
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    return cell


def create_excel(documents, pdf_map, output_path, year_label):
    """
    Erstellt die Excel-Datei im Steuerberater-Format.

    documents: Liste von Paperless-Dokumenten (API-Dicts)
    pdf_map:   {doc_id: filename_in_pdf_folder}
    output_path: Zielpfad der .xlsx-Datei
    year_label: z.B. "2024" – steht in der SUMME-Zeile
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rechnungsaufstellung"

    # ── Zeile 1: SUMME-Zeile ──────────────────────────────────────────────
    ws.row_dimensions[1].height = 20
    sum_label = ws.cell(row=1, column=1, value=f"SUMME {year_label}")
    sum_label.font = Font(bold=True, size=11)
    sum_label.alignment = Alignment(horizontal="left", vertical="center")

    # Summenformel verweist auf Spalte G (Rechnungssumme) der Tabelle
    sum_cell = ws.cell(row=1, column=7)
    sum_cell.font = Font(bold=True, size=11)
    sum_cell.alignment = Alignment(horizontal="right", vertical="center")
    sum_cell.number_format = NUMBER_FORMAT
    # Wird nach Tabellendefinition mit Formel gefüllt

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

    for i, doc in enumerate(sorted_docs):
        row = data_start_row + i
        ws.row_dimensions[row].height = 18

        # Beleg-Nr. (Paperless ASN wenn vorhanden, sonst ID)
        beleg_nr = doc.get("archive_serial_number") or doc.get("id")
        _data_cell(ws, row, 1, beleg_nr, align="center")

        # Re-Dat (created)
        created_str = doc.get("created", "")
        if created_str:
            try:
                dt = datetime.strptime(created_str[:10], "%Y-%m-%d").date()
            except ValueError:
                dt = None
        else:
            dt = None
        cell_b = _data_cell(ws, row, 2, dt, align="center",
                             number_fmt=DATE_FORMAT)

        # Zahlungsdatum – leer, manuell
        _data_cell(ws, row, 3, None, align="center",
                   number_fmt=DATE_FORMAT, bg=COLOR_EMPTY_BG)

        # Bar / Konto / KK – leer, manuell
        _data_cell(ws, row, 4, None, align="center", bg=COLOR_EMPTY_BG)

        # Beschreibung (Titel)
        _data_cell(ws, row, 5, doc.get("title", ""), align="left")

        # Kennzahl (Dokumenttyp-Name)
        doc_type_name = ""
        if doc.get("document_type_name"):
            doc_type_name = doc["document_type_name"]
        elif doc.get("document_type"):
            doc_type_name = str(doc["document_type"])
        _data_cell(ws, row, 6, doc_type_name, align="left")

        # Rechnungssumme – leer (Stufe 1), OCR-Vorschlag später (Stufe 2)
        ocr_amount = doc.get("_ocr_amount_suggestion")  # Stufe 2 Feld
        if ocr_amount is not None:
            cell_g = _data_cell(ws, row, 7, ocr_amount, align="right",
                                number_fmt=NUMBER_FORMAT, bg=COLOR_OCR_BG)
            cell_g.comment = _make_comment("OCR-Vorschlag – bitte prüfen!")
        else:
            _data_cell(ws, row, 7, None, align="right",
                       number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # Rechnungssumme inkl. Privatanteil – leer, manuell
        _data_cell(ws, row, 8, None, align="right",
                   number_fmt=NUMBER_FORMAT, bg=COLOR_EMPTY_BG)

        # Dateiname
        filename = pdf_map.get(doc.get("id"), "")
        _data_cell(ws, row, 9, filename, align="left")

    last_data_row = data_start_row + len(sorted_docs) - 1

    # ── Excel-Tabelle (Tabelle1) ──────────────────────────────────────────
    table_ref = f"A4:{get_column_letter(len(COLUMNS))}{last_data_row}"
    table = Table(displayName="Tabelle1", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    # Summenformel jetzt einsetzen (nach Tabellendefinition)
    ws["G1"] = (
        f"=SUM(G{data_start_row}:G{last_data_row})"
    )
    ws["G1"].number_format = NUMBER_FORMAT
    ws["G1"].font = Font(bold=True, size=11)
    ws["G1"].alignment = Alignment(horizontal="right", vertical="center")

    # Fenster fixieren: Header-Zeile bleibt sichtbar
    ws.freeze_panes = "A5"

    # Druckbereich
    ws.print_area = f"A1:{get_column_letter(len(COLUMNS))}{last_data_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True

    wb.save(output_path)
    return output_path


def _make_comment(text):
    """Erstellt einen openpyxl-Kommentar (nur wenn verfügbar)."""
    try:
        from openpyxl.comments import Comment
        return Comment(text, "Paperless Exporter")
    except Exception:
        return None
