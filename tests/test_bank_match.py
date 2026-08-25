"""
Tests fuer Kontoauszug-Matching (Issue #25).
"""
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault("OUTPUT_DIR", tempfile.mkdtemp())
os.environ.setdefault("PAPERLESS_URL", "http://localhost:8000")
os.environ.setdefault("PAPERLESS_TOKEN", "test-token")
os.environ.setdefault("WINDOWS_UNC_PATH", "")

from bank_csv import parse_de_amount, parse_bank_csv, csv_preview, parse_iso_date  # noqa: E402
from matching_csv import match_invoices_to_bank, amounts_match, STATUS_FOUND  # noqa: E402
from excel_export import (  # noqa: E402
    create_excel,
    read_invoices_for_matching,
    update_excel_with_bank_matches,
    write_filtered_bank_xlsx,
)
import openpyxl  # noqa: E402

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "at_giro_sample.csv")


class TestBankCsvParse:
    def test_parse_de_amount(self):
        assert parse_de_amount("-1.400,00") == -1400.0
        assert parse_de_amount("4.952,37") == 4952.37
        assert parse_de_amount("-16,05") == -16.05

    def test_parse_iso_date(self):
        assert parse_iso_date("2025-12-31").isoformat() == "2025-12-31"

    def test_preset_and_debits_only(self):
        rows = parse_bank_csv(FIXTURE, debits_only=True, only_relevant=False)
        assert all(r["amount"] is None or r["amount"] < 0 or not r["selected"] for r in rows)
        selected = [r for r in rows if r["selected"]]
        assert len(selected) == 3  # zwei -42,50 + eine -10,00; Gutschrift raus
        assert rows[0]["preset"] == "at_giro"

    def test_only_relevant(self):
        rows = parse_bank_csv(FIXTURE, debits_only=False, only_relevant=True)
        selected = [r for r in rows if r["selected"]]
        assert len(selected) == 1
        assert selected[0]["partner"] == "ACME GmbH"

    def test_preview(self):
        p = csv_preview(FIXTURE)
        assert p["preset"] == "at_giro"
        assert p["delimiter"] == ";"
        assert "Buchungsdatum" in p["fieldnames"]


class TestMatching:
    def test_amounts_match(self):
        assert amounts_match(42.5, -42.5)
        assert amounts_match(100.0, -100.01)
        assert not amounts_match(42.5, -10.0)

    def test_unique_match(self):
        from datetime import date
        invoices = [{
            "row": 5,
            "beleg_nr": 1,
            "re_dat": date(2025, 3, 10),
            "absender": "ACME GmbH",
            "beschreibung": "Rechnung Software",
            "betrag": 42.50,
        }]
        bank = parse_bank_csv(FIXTURE, debits_only=True)
        # nur erste ACME-Zeile (ohne Duplikat) fuer Eindeutigkeit
        bank = [b for b in bank if b["selected"] and "Duplikat" not in (b.get("text") or "")]
        result = match_invoices_to_bank(invoices, bank)
        assert result["stats"]["gefunden"] == 1
        assert result["matches"][0]["status"] == STATUS_FOUND
        assert result["matches"][0]["date"].isoformat() == "2025-03-15"

    def test_ambiguous_two_same_amount(self):
        from datetime import date
        invoices = [{
            "row": 5,
            "beleg_nr": 1,
            "re_dat": date(2025, 3, 10),
            "absender": "ACME",
            "beschreibung": "Rechnung",
            "betrag": 42.50,
        }]
        bank = parse_bank_csv(FIXTURE, debits_only=True)
        result = match_invoices_to_bank(invoices, bank, min_score=0.1)
        # zwei -42,50 ACME-Zeilen → mehrdeutig oder einer gefunden wenn Score klar
        assert result["stats"]["gefunden"] + result["stats"]["mehrdeutig"] >= 1


class TestExcelBankUpdate:
    def test_read_only_empty_payment_date(self, tmp_path):
        from datetime import date
        docs = [
            {"id": 1, "title": "A", "created": "2025-03-10", "archive_serial_number": 1,
             "correspondent_name": "ACME GmbH"},
            {"id": 2, "title": "B", "created": "2025-03-11", "archive_serial_number": 2,
             "correspondent_name": "Other"},
        ]
        path = str(tmp_path / "Rechnungsaufstellung_2025.xlsx")
        create_excel(docs, {}, path, "2025")
        wb = openpyxl.load_workbook(path)
        ws = wb["Rechnungsaufstellung"]
        # Betrag setzen, C leer lassen bei Zeile 5; Zeile 6 bekommt Zahlungsdatum
        ws.cell(row=5, column=8).value = 42.50
        ws.cell(row=6, column=8).value = 10.00
        ws.cell(row=6, column=3).value = date(2025, 3, 20)
        wb.save(path)
        inv = read_invoices_for_matching(path)
        assert len(inv) == 1
        assert inv[0]["beleg_nr"] == 1
        assert inv[0]["betrag"] == 42.50

    def test_update_and_filtered(self, tmp_path):
        from datetime import date
        docs = [{
            "id": 1, "title": "Software", "created": "2025-03-10",
            "archive_serial_number": 7, "correspondent_name": "ACME GmbH",
        }]
        path = str(tmp_path / "Rechnungsaufstellung_2025.xlsx")
        create_excel(docs, {}, path, "2025")
        wb = openpyxl.load_workbook(path)
        wb["Rechnungsaufstellung"].cell(row=5, column=8).value = 42.50
        wb.save(path)

        bank = parse_bank_csv(FIXTURE, debits_only=True)
        bank = [b for b in bank if b["selected"] and "Duplikat" not in (b.get("text") or "")]
        invoices = read_invoices_for_matching(path)
        result = match_invoices_to_bank(invoices, bank)
        assert result["stats"]["gefunden"] == 1
        n = update_excel_with_bank_matches(path, result)
        assert n >= 1
        wb2 = openpyxl.load_workbook(path)
        ws = wb2["Rechnungsaufstellung"]
        assert ws.cell(row=5, column=3).value is not None
        # Match-Status Spalte existiert
        headers = [ws.cell(row=4, column=c).value for c in range(1, ws.max_column + 1)]
        assert any(h and "Match-Status" in str(h) for h in headers)
        assert any(h and "Buchungstext" in str(h) for h in headers)

        out = str(tmp_path / "Kontoauszug_gefiltert_2025.xlsx")
        write_filtered_bank_xlsx(bank, result, out)
        assert os.path.exists(out)
