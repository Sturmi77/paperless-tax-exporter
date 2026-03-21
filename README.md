# paperless-tax-exporter

Steuerberater-Export für [Paperless NGX](https://docs.paperless-ngx.com/) –  
exportiert Belege als Excel-Liste im Steuerberater-Format plus PDF-Download,  
läuft als Docker Container auf dem NAS neben der Paperless-Instanz.

---

## Features

- **Web-UI** im komm|event CI für komfortable Bedienung
- **Schnellauswahl** Kalenderjahr (aktuelles Jahr + 3 Vorjahre)
- **Filter** nach Datumsbereich (von/bis) und Tags
- **Excel-Export** im Steuerberater-Format (`Tabelle1`, SUMME-Zeile, TableStyleLight1)
- **PDF-Download** aller gefilterten Belege in strukturierten Unterordner
- **Live-Log** während des Exports
- **Sicherheit**: read-only Zugriff auf Paperless, Schreibzugriff ausschließlich auf definierten Ausgabeordner
- **Stufe 2 vorbereitet**: OCR-Betragsextraktion als Vorschlag (gelb hinterlegt) geplant

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
| E | Beschreibung | Paperless Titel |
| F | Kennzahl | Paperless Dokumenttyp |
| G | Rechnungssumme | leer (Stufe 2: OCR-Vorschlag gelb) |
| H | Rechnungssumme inkl. Privatanteil | leer – manuell ausfüllen |
| I | Dateiname / Beleg | Dateiname der PDF im Belege-Ordner |

---

## Deployment

### Voraussetzungen

- Docker & Docker Compose auf dem NAS
- Paperless NGX läuft und ist erreichbar
- API-Token in Paperless generieren: **Mein Profil → Token (kreisförmiger Pfeil)**

### 1. Repository klonen

```bash
cd /volume1/docker
git clone https://github.com/Sturmi77/paperless-tax-exporter.git
cd paperless-tax-exporter
```

### 2. Token in `docker-compose.yml` eintragen

```yaml
environment:
  PAPERLESS_URL:   "http://192.168.178.115:8000"
  PAPERLESS_TOKEN: "DEIN_TOKEN_HIER"
  OUTPUT_DIR:      "/output"
```

> **Hinweis:** Den Token nie ins Git committen. Alternativ eine `.env`-Datei verwenden (liegt im `.gitignore`).

### 3. Container bauen und starten

```bash
docker compose up -d --build
```

### 4. Web-UI aufrufen

```
http://NAS-IP:5055
```

---

## Umgebungsvariablen

| Variable | Beschreibung | Standard |
|----------|-------------|---------|
| `PAPERLESS_URL` | Interne URL der Paperless-Instanz | `http://192.168.178.115:8000` |
| `PAPERLESS_TOKEN` | API-Token (read-only genügt) | – |
| `OUTPUT_DIR` | Ausgabepfad im Container | `/output` |

---

## Sicherheit

- Paperless wird **ausschließlich lesend** verwendet (nur HTTP GET)
- Schreibzugriff ist auf `OUTPUT_DIR` beschränkt und wird per `realpath`-Prüfung erzwungen
- Der Container hat **keinen Mount** auf Paperless-Daten, nur HTTP-Zugriff
- API-Token sollte in Paperless einen **read-only Benutzer** verwenden

---

## Roadmap

- [x] Stufe 1: Excel-Export + PDF-Download mit Web-UI
- [ ] Stufe 2: OCR-Betragsextraktion als Vorschlag (Spalte G, gelb hinterlegt)
- [ ] Stufe 3: Auswählbarer Ausgabepfad in der UI
- [ ] Stufe 4: Mehrere Export-Profile (verschiedene Tags / Zeiträume)

---

## Projektstruktur

```
paperless-tax-exporter/
├── app.py              # Flask-Backend, API-Endpunkte, Job-Thread
├── excel_export.py     # Excel-Generierung (openpyxl)
├── pdf_export.py       # PDF-Download aus Paperless API
├── templates/
│   └── index.html      # Web-UI
├── static/
│   ├── css/style.css   # komm|event CI
│   ├── js/app.js       # Frontend-Logik
│   └── logo.png        # komm|event Logo
├── requirements.txt
├── Dockerfile
└── docker-compose.yml
```

---

## Lizenz

Internes Projekt – keine öffentliche Lizenz.
