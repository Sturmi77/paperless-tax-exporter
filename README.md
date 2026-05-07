# paperless-tax-exporter

Steuerberater-Export für [Paperless NGX](https://docs.paperless-ngx.com/) –
exportiert Belege als Excel-Liste im Steuerberater-Format plus PDF-Download,
läuft als Docker Container auf dem NAS neben der Paperless-Instanz.

---

## Features

- **Web-UI** im komm|event CI – Teal, Helvetica Neue, Logo
- **Schnellauswahl** Kalenderjahr (aktuelles Jahr + 3 Vorjahre)
- **Filter** nach Datumsbereich, Tags und Dokumententyp (durchsuchbare Chip-Dropdowns)
- **Datumsfeld wählbar**: Belegdatum (`created`) oder Scan-Datum (`added`) – mit erklärendem Hilfetext
- **Auswählbarer Ausgabe-Unterordner**: Freitext mit Allowlist-Validierung + Live-Pfad-Preview
- **Vollständiger Export**: PDFs herunterladen + Excel + OCR-Analyse (Stufe 1+2)
- **PDFs & Excel**: Download + Aufstellung ohne OCR
- **Nur Excel**: Ohne PDF-Download (bei bereits vorhandenen PDFs)
- **OCR-Analyse**: Nachträgliche Extraktion von Absender und Betrag via Ollama (lokal, kein Cloud-LLM)
- **Nachtrag**: Neu gescannte Belege zu bestehendem Export hinzufügen ohne Neuexport
- **Überschreib-Schutz**: Modal mit drei Optionen – Nur neue hinzufügen / Alles überschreiben / Abbrechen
- **Abbrechen-Button** während OCR – Job stoppt sauber, bisherige Ergebnisse werden gespeichert
- **Live-Log** + OCR-Fortschrittsanzeige mit Ø-Zeit und ETA
- **Portable Hyperlinks**: CELL-Formel-basierte Links bleiben nach Ordner-Verschiebung gültig; Fallback auf absoluten UNC-Pfad via `HYPERLINK_MODE=unc`
- **Kopierbarer Pfad**: Optionale Spalte K mit UNC-Pfad zum Kopieren via `INCLUDE_TEXT_PATH=true`
- **Excel im komm|event CI**: Teal-Header, CI-Farben, Hyperlinks auf NAS-Freigabe
- **Sicherheit**: read-only Zugriff auf Paperless, Schreibzugriff ausschließlich auf definierten Ausgabeordner
- **Single-Worker + Threads**: Gunicorn mit 1 Worker / 4 Threads – kein Shared-State-Problem

---

## Ausgabestruktur

```
/volume1/downloads/steuerberater/
└── 2024/
    ├── Rechnungsaufstellung_2024.xlsx
    └── Belege/
        ├── 0001_McAfee_PC-Sicherheit.pdf
        ├── 0003_Amazon_Jabra_Headset.pdf
        └── …
```

Mit optionalem Unterordner (`SUBFOLDER=Privat`):

```
/volume1/downloads/steuerberater/
└── 2024/
    └── Privat/
        ├── Rechnungsaufstellung_2024.xlsx
        └── Belege/
            └── …
```

### Excel-Spalten

| Spalte | Inhalt | Quelle |
|--------|--------|--------|
| A | Beleg-Nr. | Paperless ASN / ID |
| B | Re-Dat | `created` aus Paperless |
| C | Zahlungsdatum | leer – manuell ausfüllen |
| D | Bar / Konto / KK | leer – manuell ausfüllen |
| E | Absender | Paperless Correspondent / OCR-Vorschlag (gelb) |
| F | Beschreibung | Paperless Titel |
| G | Kennzahl | Paperless Dokumenttyp |
| H | Rechnungssumme | OCR-Vorschlag (gelb) – manuell prüfen |
| I | Rechnungssumme inkl. Privatanteil | leer – manuell ausfüllen |
| J | Dateiname / Beleg | Portabler Hyperlink (CELL-Formel, bleibt nach Verschieben gültig) |
| K | UNC-Pfad | Kopierbarer absoluter Pfad (optional, `INCLUDE_TEXT_PATH=true`) |

OCR-Vorschlagswerte (Absender, Betrag) sind **gelb hinterlegt** und mit Kommentar versehen – zur manuellen Prüfung.

---

## Schnellstart (docker run)

```bash
sudo docker stop paperless-tax-exporter && sudo docker rm paperless-tax-exporter
sudo docker pull ghcr.io/sturmi77/paperless-tax-exporter:latest
sudo docker run -d \
  --name paperless-tax-exporter \
  --restart unless-stopped \
  -p 5055:5000 \
  -e PAPERLESS_URL=http://192.168.178.115:8000 \
  -e PAPERLESS_TOKEN=dein-token-hier \
  -e OLLAMA_URL=http://192.168.178.115:11434 \
  -e OLLAMA_MODEL=qwen2.5:3b \
  -e "WINDOWS_UNC_PATH=\\\\SynologyDS923\\downloads\\steuerberater" \
  -e OUTPUT_DIR=/output \
  -e TZ=Europe/Vienna \
  -v /volume1/downloads/steuerberater:/output \
  ghcr.io/sturmi77/paperless-tax-exporter:latest
```

**Web-UI:** `http://NAS-IP:5055`

---

## Update auf neue Version

```bash
sudo docker stop paperless-tax-exporter && sudo docker rm paperless-tax-exporter
sudo docker pull ghcr.io/sturmi77/paperless-tax-exporter:latest
# docker run Befehl wie oben erneut ausführen
```

---

## Umgebungsvariablen

| Variable | Beschreibung | Standard |
|----------|-------------|---------|
| `PAPERLESS_URL` | Interne URL der Paperless-Instanz | `http://192.168.178.115:8000` |
| `PAPERLESS_TOKEN` | API-Token (read-only genügt) | – |
| `OLLAMA_URL` | Ollama API URL | `http://192.168.178.115:11434` |
| `OLLAMA_MODEL` | Ollama Modell | `qwen2.5:3b` |
| `OUTPUT_DIR` | Ausgabepfad im Container | `/output` |
| `WINDOWS_UNC_PATH` | UNC-Pfad für Excel-Hyperlinks (Fallback) | `\\SynologyDS923\downloads\steuerberater` |
| `HYPERLINK_MODE` | `cell` = portable CELL-Formel (Standard); `unc` = absoluter UNC-Pfad | `cell` |
| `INCLUDE_TEXT_PATH` | `true` = Spalte K mit kopierbarem UNC-Pfad anzeigen | `false` |
| `TZ` | Zeitzone | `Europe/Vienna` |

→ Vorlage: [`.env.example`](.env.example)

---

## Deployment via docker-compose

```bash
cp .env.example .env
# .env mit eigenen Werten befüllen
docker compose up -d
```

Zum Testen eines Entwicklungs-Builds:

```bash
# In .env setzen:
IMAGE_TAG=develop
docker compose pull && docker compose up -d
```

---

## Nachtrag: Neu gescannte Belege hinzufügen

Jahresexport bereits abgeschlossen und eine vergessene Rechnung nachgescannt?

1. Gleichen Datumsbereich wie beim Erstexport wählen
2. **PDFs & Excel** oder **Vollständiger Export** starten
3. Im Dialog **„Nur neue hinzufügen"** wählen

Bereits vorhandene Belege werden übersprungen – nur neue Dokumente werden heruntergeladen und ans Excel angehängt.

---

## Portable Hyperlinks (v2.2)

Ab v2.2 verwendet Spalte J eine `CELL("filename")`-Formel statt eines statischen UNC-Pfads:

```
=HYPERLINK(LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&"Belege\datei.pdf","datei.pdf")
```

**Vorteile:**
- Excel-Datei und `Belege/`-Ordner können gemeinsam verschoben werden – Links bleiben gültig
- Kein einfrieren des Pfads beim ersten Speichern

**Voraussetzung:** Die Excel-Datei muss gespeichert sein, bevor Links funktionieren (neue ungespeicherte Dateien geben `""` zurück).

**Fallback auf v2.1-Verhalten:** `HYPERLINK_MODE=unc` in `.env` → absoluter UNC-Pfad wie bisher.

---

## Sicherheit

- Paperless wird **ausschließlich lesend** verwendet (nur HTTP GET)
- Schreibzugriff ist auf `OUTPUT_DIR` beschränkt und wird per `realpath`-Prüfung erzwungen
- Subfolder-Namen werden via Allowlist-Regex `[A-Za-z0-9_-]{1,50}` validiert (server- und clientseitig)
- `OUTPUT_DIR` wird bei Startup mit `os.path.realpath()` normalisiert (Symlink-Schutz)
- Der Container hat **keinen Mount** auf Paperless-Daten, nur HTTP-Zugriff
- Ollama läuft **lokal im Heimnetz** – kein Byte verlässt das Netzwerk
- API-Token sollte in Paperless einen **read-only Benutzer** verwenden

---

## Projektstruktur

```
paperless-tax-exporter/
├── app.py              # Flask-Backend, API-Endpunkte, Job-Thread
├── excel_export.py     # Excel-Generierung (openpyxl), komm|event CI
├── pdf_export.py       # PDF-Download aus Paperless API
├── llm_extract.py      # Ollama-Integration (Absender + Betrag)
├── templates/
│   └── index.html      # Web-UI
├── static/
│   ├── css/style.css   # komm|event CI
│   ├── js/app.js       # Frontend-Logik
│   └── logo.png        # komm|event Logo
├── tests/
│   ├── test_excel.py   # Excel-Export Tests (49 Tests)
│   └── test_security.py # Sicherheits- und Allowlist-Tests
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── .env.example
```

---

## Changelog

### v2.2.0 (2026-04-19)
- **Issue #7** – Dokumententyp als Filterkriterium auswählbar (Chip-Dropdown, analog Tags)
- **Issue #8** – Portable Hyperlinks via `CELL("filename")`-Formel; Fallback `HYPERLINK_MODE=unc`; optionale Spalte K (`INCLUDE_TEXT_PATH`)
- **Issue #9** – Auswählbarer Ausgabe-Unterordner mit Allowlist-Validierung und Pfad-Preview
- **Issue #10** – Rechnungsdatum/Scandatum Toggle: Labels, Hilfetext, Segmented-Control-Design
- **Refactoring** – `createChipDropdown()` Factory; `subfolder`-Schnittstelle in allen Excel-Funktionen; 49 automatisierte Tests

### v2.1.0
- Nachtrag-Funktion (neue Belege anhängen ohne Neuexport)
- Überschreib-Schutz mit Modal
- Abbrechen-Button während OCR
- Live-Log + ETA-Anzeige

---

## Roadmap

### v2.3 – GUI-Verbesserungen *(geplant)*

| # | Feature |
|---|---------|
| [#12](https://github.com/Sturmi77/paperless-tax-exporter/issues/12) | Logo + Favicon, konfigurierbarer App-Name (`APP_TITLE`) |
| [#12](https://github.com/Sturmi77/paperless-tax-exporter/issues/12) | Tags- und Dokumententyp-Filter nebeneinander (Responsive Layout) |
| [#12](https://github.com/Sturmi77/paperless-tax-exporter/issues/12) | Subfolder-Picker als Modal (Ordnerstruktur auswählen + neuen Ordner anlegen) |
| [#12](https://github.com/Sturmi77/paperless-tax-exporter/issues/12) | Hinweistext in Kopfzeile inline, aktuelles Kalenderjahr vorausgewählt |
| [#19](https://github.com/Sturmi77/paperless-tax-exporter/issues/19) | Fortschrittsanzeige (Ø-Zeit + ETA) in allen Export-Stufen |

### v3.0 – Rudimentäre Buchhaltung *(Konzeptphase)*

Aufbauend auf dem bestehenden Eingangsrechnungs-Export entsteht schrittweise eine
vollständige Buchhaltungsbasis – lokal, ohne Cloud-Dienste.

#### Milestone 1 – Ausgangsrechnungen ([#20](https://github.com/Sturmi77/paperless-tax-exporter/issues/20))

- Ausgangsrechnungen aus Paperless exportieren (eigener Dokumententyp-Filter, aufbauend auf Issue #7)
- LLM-gestützte Extraktion von Betrag, Rechnungsnummer und Empfänger via Ollama
- Separates Excel `Ausgangsrechnungen_{Jahr}.xlsx` im komm|event CI
- Spalten: Rechnungsnummer, Datum, Fälligkeitsdatum, Empfänger, Betrag netto/brutto, Zahlungsstatus

#### Milestone 2 – Monatliche G+V-Berichte ([#20](https://github.com/Sturmi77/paperless-tax-exporter/issues/20))

- Gegenüberstellung Einnahmen (Ausgangsrechnungen) vs. Ausgaben (Eingangsrechnungen)
- Gruppierung nach Dokumententyp / Kennzahl (Spalte G)
- Filterbarer Zeitraum: Monat, Quartal, Jahr
- Ausgabe als separates Excel-Sheet oder PDF

#### Milestone 3 – Mahnwesen ([#20](https://github.com/Sturmi77/paperless-tax-exporter/issues/20) + [#11](https://github.com/Sturmi77/paperless-tax-exporter/issues/11))

- Aufbauend auf Issue #11 (Buchungsbestätigungen per LLM zuordnen):
  Ausgangsrechnung + Fälligkeitsdatum bekannt → fehlende Zahlung = offener Posten
- Export einer Offene-Posten-Liste
- Optional: Mahnschreiben-Vorlage aus unbezahlten Rechnungen befüllen


---

## Lizenz

Internes Projekt – keine öffentliche Lizenz.
