"""
Backend-Tests fuer die Subfolder-Picker API (Issue #12, Schritt 4).
Testet GET /api/subfolders und POST /api/subfolders.
"""
import pytest
import sys
import os
import tempfile
import json

# Projektpfad
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


# ---------------------------------------------------------------------------
# Flask-Test-Client Setup
# ---------------------------------------------------------------------------
@pytest.fixture()
def app_with_output(tmp_path):
    """
    Flask-App mit temporaerem OUTPUT_DIR starten.
    Nutzt Environment-Variable damit app.py den Pfad setzt.
    """
    os.environ["OUTPUT_DIR"]          = str(tmp_path)
    os.environ["PAPERLESS_URL"]       = "http://localhost:8000"
    os.environ["PAPERLESS_TOKEN"]     = "test-token-123"
    os.environ["WINDOWS_UNC_PATH"]    = ""

    import importlib
    import app as app_module
    importlib.reload(app_module)

    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as client:
        yield client, tmp_path

    # Cleanup env
    for key in ["OUTPUT_DIR", "PAPERLESS_URL", "PAPERLESS_TOKEN", "WINDOWS_UNC_PATH"]:
        os.environ.pop(key, None)


# ---------------------------------------------------------------------------
# Tests: GET /api/subfolders
# ---------------------------------------------------------------------------
class TestGetSubfolders:
    def test_empty_output_dir_returns_empty_list(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        assert data["subfolders"] == []

    def test_existing_valid_dirs_returned(self, app_with_output):
        client, tmp_path = app_with_output
        (tmp_path / "Archiv").mkdir()
        (tmp_path / "2024-Q4").mkdir()
        (tmp_path / "steuer_2024").mkdir()
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        assert sorted(data["subfolders"]) == ["2024-Q4", "Archiv", "steuer_2024"]

    def test_invalid_dirs_excluded(self, app_with_output):
        """Ordner mit Leerzeichen, Sonderzeichen oder > 50 Zeichen werden gefiltert."""
        client, tmp_path = app_with_output
        (tmp_path / "Mein Ordner").mkdir()       # Leerzeichen
        (tmp_path / "folder.test").mkdir()        # Punkt
        (tmp_path / "valid-folder").mkdir()       # gueltig
        (tmp_path / ("A" * 51)).mkdir()           # zu lang
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        # Nur "valid-folder" bleibt
        assert data["subfolders"] == ["valid-folder"]

    def test_files_excluded(self, app_with_output):
        """Regulaere Dateien werden nicht als Unterordner aufgelistet."""
        client, tmp_path = app_with_output
        (tmp_path / "Archiv").mkdir()
        (tmp_path / "report.xlsx").write_text("dummy")
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        assert data["subfolders"] == ["Archiv"]

    def test_sorted_alphabetically(self, app_with_output):
        client, tmp_path = app_with_output
        (tmp_path / "Zebra").mkdir()
        (tmp_path / "Alpha").mkdir()
        (tmp_path / "Mitte").mkdir()
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        assert data["subfolders"] == ["Alpha", "Mitte", "Zebra"]

    def test_response_has_subfolders_key(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.get("/api/subfolders")
        assert res.status_code == 200
        data = json.loads(res.data)
        assert "subfolders" in data


# ---------------------------------------------------------------------------
# Tests: POST /api/subfolders
# ---------------------------------------------------------------------------
class TestPostSubfolders:
    def test_create_valid_folder(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "Archiv"}),
            content_type="application/json",
        )
        assert res.status_code == 201
        data = json.loads(res.data)
        assert data["created"] == "Archiv"
        assert (tmp_path / "Archiv").is_dir()

    def test_create_folder_with_dash_and_underscore(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "steuer_2024-Q4"}),
            content_type="application/json",
        )
        assert res.status_code == 201
        assert (tmp_path / "steuer_2024-Q4").is_dir()

    def test_create_folder_idempotent(self, app_with_output):
        """Bereits existierender Ordner wird akzeptiert (exist_ok=True)."""
        client, tmp_path = app_with_output
        (tmp_path / "Archiv").mkdir()
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "Archiv"}),
            content_type="application/json",
        )
        assert res.status_code == 201

    def test_empty_name_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": ""}),
            content_type="application/json",
        )
        assert res.status_code == 400
        data = json.loads(res.data)
        assert "error" in data

    def test_missing_name_key_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_path_traversal_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "../etc"}),
            content_type="application/json",
        )
        assert res.status_code == 400
        data = json.loads(res.data)
        assert "error" in data

    def test_slash_in_name_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "folder/sub"}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_space_in_name_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "Mein Ordner"}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_too_long_name_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "A" * 51}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_absolute_path_rejected(self, app_with_output):
        client, tmp_path = app_with_output
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": "/etc/passwd"}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_created_folder_appears_in_list(self, app_with_output):
        """Neu angelegter Ordner erscheint direkt in GET /api/subfolders."""
        client, tmp_path = app_with_output
        client.post(
            "/api/subfolders",
            data=json.dumps({"name": "Neu2024"}),
            content_type="application/json",
        )
        res = client.get("/api/subfolders")
        data = json.loads(res.data)
        assert "Neu2024" in data["subfolders"]

    def test_max_length_name_accepted(self, app_with_output):
        client, tmp_path = app_with_output
        name = "A" * 50
        res = client.post(
            "/api/subfolders",
            data=json.dumps({"name": name}),
            content_type="application/json",
        )
        assert res.status_code == 201
        assert (tmp_path / name).is_dir()
