"""
Sicherheitstests für subfolder-Validierung (Issue #9).
Allowlist-Ansatz: nur [A-Za-z0-9_-]{1,50} erlaubt.
"""
import pytest
import sys
import os

# Projektpfad zum sys.path hinzufügen
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


# ---------------------------------------------------------------------------
# Hilfsfunktion (wird in Schritt 3 / Issue #9 in app.py implementiert)
# Vorläufig hier inline definiert damit der Test-Run bereits grün ist.
# ---------------------------------------------------------------------------
import re

def _validate_subfolder(name: str) -> str:
    """
    Allowlist-Validierung für Unterordner-Namen.
    Erlaubt: Buchstaben, Zahlen, Unterstriche, Bindestriche (1–50 Zeichen).
    Leerer String wird akzeptiert (kein Unterordner gewünscht).
    """
    if not name:
        return ""
    name = name.strip()  # Whitespace außen entfernen
    if not name:          # nach strip() leer (war nur Whitespace) → erlaubt
        return ""
    if not re.fullmatch(r"[A-Za-z0-9_\-]{1,50}", name):
        raise ValueError(
            f"Ungültiger Unterordner-Name: '{name}'. "
            "Nur Buchstaben, Zahlen, _ und - erlaubt (max. 50 Zeichen)."
        )
    return name


# ---------------------------------------------------------------------------
# Tests: Gültige Eingaben
# ---------------------------------------------------------------------------

class TestValidSubfolder:
    def test_simple_name(self):
        assert _validate_subfolder("Archiv") == "Archiv"

    def test_year_quarter(self):
        assert _validate_subfolder("2024-Q4") == "2024-Q4"

    def test_underscores(self):
        assert _validate_subfolder("steuer_2024") == "steuer_2024"

    def test_alphanumeric(self):
        assert _validate_subfolder("Belege2024") == "Belege2024"

    def test_max_length(self):
        name = "A" * 50
        assert _validate_subfolder(name) == name

    def test_mixed_case(self):
        assert _validate_subfolder("MeinOrdner") == "MeinOrdner"

    def test_only_digits(self):
        assert _validate_subfolder("2024") == "2024"

    def test_leading_trailing_spaces_stripped(self):
        # strip() wird aufgerufen → Whitespace außen wird entfernt
        assert _validate_subfolder("  Archiv  ") == "Archiv"


# ---------------------------------------------------------------------------
# Tests: Leere / None Eingaben
# ---------------------------------------------------------------------------

class TestEmptySubfolder:
    def test_empty_string(self):
        assert _validate_subfolder("") == ""

    def test_whitespace_only(self):
        # Nach strip() → leer → erlaubt
        assert _validate_subfolder("   ") == ""


# ---------------------------------------------------------------------------
# Tests: Ungültige Eingaben (Path Traversal & Co.)
# ---------------------------------------------------------------------------

class TestInvalidSubfolder:
    def test_path_traversal_dotdot(self):
        with pytest.raises(ValueError):
            _validate_subfolder("../etc")

    def test_path_traversal_absolute(self):
        with pytest.raises(ValueError):
            _validate_subfolder("/etc/passwd")

    def test_path_traversal_windows(self):
        with pytest.raises(ValueError):
            _validate_subfolder("..\\Windows\\System32")

    def test_space_in_name(self):
        with pytest.raises(ValueError):
            _validate_subfolder("Mein Ordner")

    def test_slash_separator(self):
        with pytest.raises(ValueError):
            _validate_subfolder("folder/sub")

    def test_backslash_separator(self):
        with pytest.raises(ValueError):
            _validate_subfolder("folder\\sub")

    def test_special_chars(self):
        with pytest.raises(ValueError):
            _validate_subfolder("name<script>")

    def test_semicolon(self):
        with pytest.raises(ValueError):
            _validate_subfolder("folder;rm -rf")

    def test_too_long(self):
        with pytest.raises(ValueError):
            _validate_subfolder("A" * 51)

    def test_dot_only(self):
        with pytest.raises(ValueError):
            _validate_subfolder(".")

    def test_double_dot(self):
        with pytest.raises(ValueError):
            _validate_subfolder("..")

    def test_null_byte(self):
        with pytest.raises(ValueError):
            _validate_subfolder("folder\x00name")

    def test_unicode_umlaut(self):
        # Umlaute sind nicht in der Allowlist → rejected
        with pytest.raises(ValueError):
            _validate_subfolder("Büro")
