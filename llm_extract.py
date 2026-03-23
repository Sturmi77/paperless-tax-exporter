"""
llm_extract.py – Extraktion von Absender und Rechnungsbetrag via Ollama (lokal).

Strategie:
  - Sendet den OCR-Text (content-Feld aus Paperless) an Ollama
  - Erwartet strukturiertes JSON zurück: { "absender": "...", "betrag": 123.45 }
  - Kein Byte verlässt das Heimnetz
"""

import os
import re
import json
import requests

OLLAMA_URL   = os.environ.get("OLLAMA_URL",   "http://192.168.178.115:11434")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "qwen2.5:3b")

PROMPT_TEMPLATE = """Du analysierst den OCR-Text einer österreichischen oder deutschen Rechnung.
Extrahiere genau zwei Informationen und antworte NUR mit gültigem JSON, ohne Erklärungen.

JSON-Format:
{{
  "absender": "Name des Rechnungsstellers (Firma oder Person)",
  "betrag": 123.45
}}

Regeln:
- "absender": Vollständiger Name des Rechnungsausstellers. Falls nicht eindeutig: null
- "betrag": Gesamtbetrag inkl. MwSt als Dezimalzahl (Punkt als Trennzeichen). Falls nicht gefunden: null
- Nur reines JSON zurückgeben, kein Markdown, keine Erklärungen

OCR-Text:
{text}"""


def _extract_relevant_text(content: str, max_chars: int = 1000) -> str:
    """
    Sendet nur den relevanten Teil des OCR-Textes ans LLM.
    Rechnungssummen und Absender stehen meist am Anfang und Ende.
    """
    if not content:
        return ""
    lines = [l.strip() for l in content.splitlines() if l.strip()]
    # Erste 30 + letzte 30 Zeilen (Briefkopf + Summenbereich)
    relevant = lines[:30] + (["..."] if len(lines) > 60 else []) + lines[-30:]
    text = "\n".join(relevant)
    return text[:max_chars]


def extract_from_ocr(content: str, timeout: int = 180) -> dict:
    """
    Sendet OCR-Text an Ollama und gibt dict zurück:
    {
        "absender": str | None,
        "betrag":   float | None,
        "error":    str | None   (nur bei Fehler)
    }
    """
    result = {"absender": None, "betrag": None, "error": None}

    if not content or not content.strip():
        result["error"] = "Kein OCR-Text verfügbar"
        return result

    text = _extract_relevant_text(content)
    prompt = PROMPT_TEMPLATE.format(text=text)

    try:
        resp = requests.post(
            f"{OLLAMA_URL.rstrip('/')}/api/generate",
            json={
                "model":  OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False,
                "format": "json",   # Ollama structured output
                "options": {
                    "temperature": 0.0,   # deterministisch
                    "num_predict": 150,   # kurze Antwort reicht
                    "num_ctx": 1024,      # kleiner Context = schneller
                }
            },
            timeout=timeout,
        )
        resp.raise_for_status()
        raw = resp.json().get("response", "")

        # JSON aus Antwort extrahieren (manchmal Markdown-Wrapper)
        json_match = re.search(r'\{.*?\}', raw, re.DOTALL)
        if not json_match:
            result["error"] = f"Kein JSON in Antwort: {raw[:100]}"
            return result

        parsed = json.loads(json_match.group())
        result["absender"] = parsed.get("absender") or None
        raw_betrag = parsed.get("betrag")

        # Betrag normalisieren (String "1.234,56" → float 1234.56)
        if raw_betrag is not None:
            result["betrag"] = _normalize_amount(raw_betrag)

    except requests.exceptions.Timeout:
        result["error"] = f"Ollama Timeout nach {timeout}s"
    except requests.exceptions.ConnectionError:
        result["error"] = f"Ollama nicht erreichbar ({OLLAMA_URL})"
    except json.JSONDecodeError as e:
        result["error"] = f"JSON-Parsing-Fehler: {e}"
    except Exception as e:
        result["error"] = f"Unbekannter Fehler: {e}"

    return result


def _normalize_amount(value) -> float | None:
    """Normalisiert Betragsformate zu float."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        # "1.234,56" → "1234.56"
        v = value.strip().replace("€", "").replace(" ", "")
        if "," in v and "." in v:
            # 1.234,56 Format
            v = v.replace(".", "").replace(",", ".")
        elif "," in v:
            # 1234,56 Format
            v = v.replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return None
    return None


def check_ollama_available() -> tuple[bool, str]:
    """Prüft ob Ollama erreichbar ist und das Modell verfügbar ist."""
    try:
        resp = requests.get(f"{OLLAMA_URL.rstrip('/')}/api/tags", timeout=5)
        resp.raise_for_status()
        models = [m["name"] for m in resp.json().get("models", [])]
        model_base = OLLAMA_MODEL.split(":")[0]
        available = any(m.startswith(model_base) for m in models)
        if not available:
            return False, f"Modell '{OLLAMA_MODEL}' nicht gefunden. Verfügbar: {models}"
        return True, f"Ollama OK – Modell: {OLLAMA_MODEL}"
    except requests.exceptions.ConnectionError:
        return False, f"Ollama nicht erreichbar ({OLLAMA_URL})"
    except Exception as e:
        return False, str(e)
