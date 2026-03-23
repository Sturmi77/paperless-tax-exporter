# paperless-tax-exporter

Steuerberater-Export für [Paperless NGX](https://docs.paperless-ngx.com/) –
exportiert Belege als Excel-Liste im Steuerberater-Format plus PDF-Download,
läuft als Docker Container auf dem NAS neben der Paperless-Instanz.

---

## Features

- **Web-UI** im komm|event CI – Teal, Helvetica Neue, Logo
- **Schnellauswahl** Kalenderjahr (aktuelles Jahr + 3 Vorjahre)
- **Filter** nach Datumsbereich und Tags (durchsuchbares Chip-Dropdown)
- **Datumsfeld wählbar**: Belegdatum (`created`) oder Scan-Datum (`added`)
- **Vollständiger Export**: PDFs herunterladen + Excel + OCR-Analyse (Stufe 1+2)
- **PDFs & Excel**: Download + Aufstellung ohne OCR
- **Nur Excel**: Ohne PDF-Download (bei bereits vorhandenen PDFs)
- **OCR-Analyse**: Nachträgliche Extraktion von Absender und Betrag via Ollama (lokal, kein Cloud-LLM)
- **Nachtrag**: Neu gescannte Belege zu bestehendem Export hinzufügen ohne Neuexport
- **Überschreib-Schutz**: Modal mit drei Optionen – Nur neue hinzufügen / Alles überschreiben / Abbrechen
- **Abbrechen-Button** während OCR – Job stoppt sauber, bisherige Ergebnisse werden gespeichert
- **Live-Log** + OCR-Fortschrittsanzeige mit Ø-Zeit und ETA
- **Excel im komm|event CI**: Teal-Header, CI-Farben, Hyperlinks auf NAS-Freigabe (UNC-Pfad)
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
| J | Dateiname / Beleg | Hyperlink auf NAS-Freigabe (UNC-Pfad) |

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
| `WINDOWS_UNC_PATH` | UNC-Pfad für Excel-Hyperlinks | `\\SynologyDS923\downloads\steuerberater` |
| `TZ` | Zeitzone | `Europe/Vienna` |

→ Vorlage: [`.env.example`](.env.example)

---

## Deployment via docker-compose

```bash
cp .env.example .env
# .env mit eigenen Werten befüllen
docker compose up -d
```

---

## Nachtrag: Neu gescannte Belege hinzufügen

Jahresexport bereits abgeschlossen und eine vergessene Rechnung nachgescannt?

1. Gleichen Datumsbereich wie beim Erstexport wählen
2. **PDFs & Excel** oder **Vollständiger Export** starten
3. Im Dialog **„Nur neue hinzufügen"** wählen

Bereits vorhandene Belege werden übersprungen – nur neue Dokumente werden heruntergeladen und ans Excel angehängt.

---

## Sicherheit

- Paperless wird **ausschließlich lesend** verwendet (nur HTTP GET)
- Schreibzugriff ist auf `OUTPUT_DIR` beschränkt und wird per `realpath`-Prüfung erzwungen
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
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── .env.example
```

---

## Lizenz

Internes Projekt – keine öffentliche Lizenz.
