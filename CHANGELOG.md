# Changelog

Alle nennenswerten Änderungen werden hier dokumentiert.
Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.0.0/).

---

## [2.3.0] – 2026-05-07

### Hinzugefügt

#### Subfolder-Picker Modal (Issue #12, Schritt 4)
- `GET /api/subfolders` – listet alle validen Unterverzeichnisse von `OUTPUT_DIR` (Allowlist `[A-Za-z0-9_-]{1,50}`, alphabetisch, keine Symlinks)
- `POST /api/subfolders` – legt neuen Unterordner an mit `_validate_subfolder` + `_assert_output_path` (kein Path-Traversal möglich)
- Picker-Button (`aria-haspopup=dialog`) ersetzt Freitext-Eingabe
- Modal: Ordnerliste (`role=listbox`), Suchfeld (Echtzeit-Filter), Neuer-Ordner-Bereich
- „Kein Unterordner" Reset-Option immer oben in der Liste
- Focus-Trap: Escape, Backdrop-Klick, X-Button; Fokus-Rückgabe an Picker-Button
- Fallback: Hinweis im Modal wenn API nicht erreichbar, kein JS-Fehler

#### JS/UX-Fixes (Issue #12, Schritt 3)
- **A4** Pill-Toggle (Jahr/Tag/Datum) mit Arrow-Key-Navigation (ARIA `role=radiogroup`, roving tabindex)
- **F3** `aria-disabled` auf Export-Buttons synchron mit Formular-Validierung; `selected-info` zeigt Initialtext „Datumsbereich wählen, um Export zu starten."
- **F4** Subfolder-Bereich als `<fieldset>`/`<legend>` für semantisches Grouping
- **S4** Globaler Error-Banner (`role=alert`) bei Paperless-API-Verbindungsproblemen mit „Verbindung wiederherstellen"-Button
- **P3** Phasen-Label in `progress-title` bei Stage-Wechsel (data-phase verhindert unnötige DOM-Updates)
- **P4** Fokus auf `btn-cancel` beim ersten Einblenden des Abbrechen-Buttons

#### Layout & UX (Issue #12, Schritte 1+2)
- Logo (`kommevent.at`) + Favicon im Browser-Tab
- `APP_TITLE` ENV-Variable für konfigurierbaren App-Namen (Standard: `Steuerberater Export`)
- Aktuelles Kalenderjahr automatisch vorausgewählt (`buildYearButtons()`)
- `aria-pressed` auf Year-Buttons
- Tags- und Dokumententyp-Filter nebeneinander (CSS Grid `1fr 1fr`, responsive auf Mobile untereinander)
- Hinweistext inline neben „Export konfigurieren" (`.card-title-row`, gedimmt)
- Export-Button Redesign: Primärbutton volle Breite + Sekundär-Reihe
- WCAG-konformer Fokus-Ring (`outline: 2px solid var(--primary)`)
- Touch-Targets `min-height: 44px` auf Mobile
- Connection-Badge mit SVG-Icon
- Modal: `max-height: 90vh` + scrollbar, `flex-wrap` auf Header unter 360px

### Behoben

- **Paginierungsbug** – `get_documents()` folgte `page=N` statt `next`-URL; Paperless verwendet cursor-basierte Paginierung. Fix: erste Seite mit params, Folgeseiten via `next`-URL ohne eigene params. Dokumente außerhalb des Datumsbereichs erschienen nicht mehr.
- **Webkit SyntaxError (mehrzeilige Template-Literals)** – Alle mehrzeiligen Backtick-Strings in `innerHTML`-Zuweisungen wurden auf einzeilige Strings umgestellt. Betroffen: `banner.innerHTML`, `cancelBtn.innerHTML`, `btn.innerHTML`, `opt.innerHTML`, `noneItem.innerHTML`, `item.innerHTML`. Ursache: Synology DSM WebKit-Browser akzeptiert keine Zeilenumbrüche innerhalb von Template-Literals in `innerHTML`.
- **Webkit SyntaxError (typografische Anführungszeichen)** – Unicode `„` (U+201E) in `empty.textContent`-Zuweisung durch ASCII-Äquivalent ersetzt. Ursache: Python-Editor fügte automatisch typografische Anführungszeichen ein, die der WebKit-Parser als unerwartetes String-Ende interpretierte.

### Tests

- 18 neue Tests in `tests/test_subfolders_api.py` (GET: 6, POST: 12)
- Gesamt: **67/67** (vorher: 49)
- Abgedeckt: Path-Traversal, Allowlist-Validierung, Idempotenz, Sortierung, Datei-Filterung, API-Fehlerbehandlung

---

## [2.2.0] – 2026-04-19

### Hinzugefügt

- **Issue #7** – Dokumententyp als Filterkriterium (Chip-Dropdown, analog Tags)
- **Issue #8** – Portable Hyperlinks via `CELL("filename")`-Formel; Fallback `HYPERLINK_MODE=unc`; optionale Spalte K `INCLUDE_TEXT_PATH=true`
- **Issue #9** – Auswählbarer Ausgabe-Unterordner mit Allowlist-Validierung `[A-Za-z0-9_-]{1,50}` und Pfad-Preview
- **Issue #10** – Rechnungsdatum/Scandatum Toggle: Labels, Hilfetext, Segmented-Control-Design
- `createChipDropdown()` Factory-Funktion für wiederverwendbare Chip-Dropdowns
- `subfolder`-Schnittstelle in allen Excel-Funktionen (`create_excel`, `append_to_excel`)

### Tests

- 49 automatisierte Tests (Excel-Export, UNC-Pfad-Logik, CELL-Formel, Allowlist-Validierung)

---

## [2.1.0]

### Hinzugefügt

- Nachtrag-Funktion: neu gescannte Belege ohne Neuexport anhängen
- Überschreib-Schutz: Modal mit drei Optionen (Nur neue / Alles überschreiben / Abbrechen)
- Abbrechen-Button während OCR – Job stoppt sauber
- Live-Log + OCR-Fortschrittsanzeige mit Ø-Zeit und ETA

---

## Bekannte Einschränkungen / Webkit-Kompatibilität

Der Container läuft auf Synology DSM. Beim Entwickeln mit modernen IDEs/Editoren sind folgende Punkte zu beachten:

1. **Mehrzeilige Template-Literals in `innerHTML`** – Webkit (DSM-Browser) akzeptiert keine Zeilenumbrüche innerhalb von Template-Literals die als `innerHTML` gesetzt werden. Alle solchen Strings müssen einzeilig sein.
2. **Typografische Anführungszeichen** – Editoren ersetzen manchmal `"` automatisch durch `„`/`"` (U+201E/U+201C). Diese Zeichen sind in JS-String-Literalen auf älteren Webkit-Versionen nicht erlaubt.
3. **Prüfung vor Commit** – Mit folgendem Python-Snippet können beide Fehlerbilder geprüft werden:

```python
content = open('static/js/app.js').read()
lines = content.split('\n')

# Mehrzeilige Template-Literals
in_t = False
for i, line in enumerate(lines, 1):
    bt = line.find('`')
    if not in_t and bt >= 0 and '`' not in line[bt+1:]:
        in_t = True; start = i
    elif in_t and '`' in line:
        in_t = False
        print(f"WARN: Mehrzeiliges Template-Literal Zeilen {start}-{i}")

# Typografische Sonderzeichen
for i, line in enumerate(lines, 1):
    for ch in ['\u201e', '\u201c', '\u201d']:
        if ch in line:
            print(f"WARN: Typograf. Anführungszeichen U+{ord(ch):04X} in Zeile {i}")
```
