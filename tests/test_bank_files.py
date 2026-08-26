"""
Tests fuer Dateiauswahl CSV/Excel unter OUTPUT_DIR (Issue #28 / #29).
"""
import json
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture()
def app_with_output(tmp_path):
    os.environ["OUTPUT_DIR"] = str(tmp_path)
    os.environ["PAPERLESS_URL"] = "http://localhost:8000"
    os.environ["PAPERLESS_TOKEN"] = "test-token-123"
    os.environ["WINDOWS_UNC_PATH"] = ""

    import importlib
    import app as app_module
    importlib.reload(app_module)

    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as client:
        yield client, tmp_path, app_module

    for key in ["OUTPUT_DIR", "PAPERLESS_URL", "PAPERLESS_TOKEN", "WINDOWS_UNC_PATH"]:
        os.environ.pop(key, None)


class TestListFiles:
    def test_list_csv_and_xlsx(self, app_with_output):
        client, tmp_path, _ = app_with_output
        year = tmp_path / "2025"
        year.mkdir()
        (year / "Rechnungsaufstellung_2025.xlsx").write_bytes(b"PK\x03\x04")
        (year / "Bearbeitet_Steuer.xlsx").write_bytes(b"PK\x03\x04")
        (tmp_path / "umsaetze.csv").write_text("a;b\n1;2\n", encoding="utf-8")

        res = client.get("/api/bank-csv/files?type=xlsx")
        assert res.status_code == 200
        data = json.loads(res.data)
        paths = {f["rel_path"] for f in data["files"]}
        assert "2025/Rechnungsaufstellung_2025.xlsx" in paths
        assert "2025/Bearbeitet_Steuer.xlsx" in paths

        res2 = client.get("/api/bank-csv/files?type=csv")
        assert res2.status_code == 200
        data2 = json.loads(res2.data)
        assert any(f["rel_path"] == "umsaetze.csv" for f in data2["files"])

    def test_rejects_invalid_type(self, app_with_output):
        client, _, _ = app_with_output
        res = client.get("/api/bank-csv/files?type=exe")
        assert res.status_code == 400


class TestResolvePaths:
    def test_normalize_rejects_traversal(self, app_with_output):
        _, _, app_module = app_with_output
        with pytest.raises(ValueError):
            app_module._normalize_rel_path("../etc/passwd")
        with pytest.raises(ValueError):
            app_module._normalize_rel_path("2025/../../x.csv")

    def test_preview_with_csv_rel_path(self, app_with_output):
        client, tmp_path, _ = app_with_output
        fixture = os.path.join(
            os.path.dirname(__file__), "fixtures", "at_giro_sample.csv"
        )
        target = tmp_path / "konto.csv"
        target.write_text(open(fixture, encoding="utf-8").read(), encoding="utf-8")

        res = client.post(
            "/api/bank-csv/preview",
            json={"csv_rel_path": "konto.csv"},
        )
        assert res.status_code == 200
        data = json.loads(res.data)
        assert data["preset"] == "at_giro"
        assert data["csv_rel_path"] == "konto.csv"

    def test_match_with_excel_and_csv_rel(self, app_with_output):
        client, tmp_path, app_module = app_with_output
        from excel_export import create_excel
        from datetime import date
        import openpyxl

        year = tmp_path / "2025"
        year.mkdir()
        excel = year / "Mein_Export.xlsx"
        docs = [{
            "id": 1, "title": "Software", "created": "2025-03-10",
            "archive_serial_number": 7, "correspondent_name": "ACME GmbH",
        }]
        create_excel(docs, {}, str(excel), "2025")
        wb = openpyxl.load_workbook(str(excel))
        wb["Rechnungsaufstellung"].cell(row=5, column=8).value = 42.50
        wb.save(str(excel))

        fixture = os.path.join(
            os.path.dirname(__file__), "fixtures", "at_giro_sample.csv"
        )
        csv_path = tmp_path / "konto.csv"
        # Nur eine -42,50-Zeile (ohne Duplikat) fuer eindeutigen Match
        lines = open(fixture, encoding="utf-8").read().splitlines()
        header, rows = lines[0], lines[1:]
        rows = [r for r in rows if "Duplikat" not in r]
        csv_path.write_text(header + "\n" + "\n".join(rows) + "\n", encoding="utf-8")

        res = client.post(
            "/api/bank-csv/match",
            json={
                "excel_rel_path": "2025/Mein_Export.xlsx",
                "csv_rel_path": "konto.csv",
                "dry_run": True,
                "debits_only": True,
            },
        )
        assert res.status_code == 200, res.data
        data = json.loads(res.data)
        assert data["excel_rel_path"] == "2025/Mein_Export.xlsx"
        assert data["stats"]["gefunden"] >= 1
