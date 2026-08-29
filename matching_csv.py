"""
Matching Kontoauszug-Buchungen ↔ Rechnungszeilen (Issue #25 / #34).

Härte-Regeln:
  1. Betrag (± Toleranz)
  2. Datum-Fenster (Rechnungsdatum … + max_days)
  3. Fuzzy-Text mit Mindestscore (kein Blind-Match nur über Betrag)
  4. 1:1-Lock: jede Bankzeile höchstens einer Rechnung
  5. ≥2 Kandidaten mit Score ≥ min_score → immer mehrdeutig

Nur Rechnungszeilen mit leerem Zahlungsdatum (Spalte C).
"""

from __future__ import annotations

import re
from datetime import date, timedelta
from difflib import SequenceMatcher
from typing import Optional


STATUS_OPEN = "offen"
STATUS_FOUND = "gefunden"
STATUS_AMBIGUOUS = "mehrdeutig"
STATUS_NOT_FOUND = "nicht gefunden"
STATUS_NO_AMOUNT = "kein Betrag"


def _norm_text(s: str) -> str:
    if not s:
        return ""
    s = s.lower()
    s = re.sub(r"[^\w\säöüÄÖÜß]+", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def text_score(a: str, b: str) -> float:
    """Ähnlichkeit 0..1 zwischen zwei Freitexten."""
    na, nb = _norm_text(a), _norm_text(b)
    if not na or not nb:
        return 0.0
    if na in nb or nb in na:
        return 0.95
    ta, tb = set(na.split()), set(nb.split())
    if ta and tb:
        overlap = len(ta & tb) / max(len(ta), len(tb))
    else:
        overlap = 0.0
    seq = SequenceMatcher(None, na, nb).ratio()
    return max(overlap, seq)


def amounts_match(invoice_amount: float, bank_amount: float, tol_abs: float = 0.02,
                  tol_pct: float = 0.01) -> bool:
    """Vergleicht Rechnungsbetrag mit |Bankbetrag|."""
    if invoice_amount is None or bank_amount is None:
        return False
    bank_abs = abs(bank_amount)
    inv = abs(float(invoice_amount))
    diff = abs(inv - bank_abs)
    return diff <= max(tol_abs, inv * tol_pct)


def date_in_window(invoice_date: Optional[date], bank_date: Optional[date],
                   max_days: int = 60) -> bool:
    if not invoice_date or not bank_date:
        return True  # kein Datum → nicht disqualifizieren
    if bank_date < invoice_date:
        return False
    return bank_date <= invoice_date + timedelta(days=max_days)


def score_candidate(invoice: dict, bank: dict) -> float:
    """Gesamtscore 0..1 für ein Invoice/Bank-Paar (Betrag bereits geprüft)."""
    parts = []
    inv_text = " ".join(filter(None, [
        str(invoice.get("absender") or ""),
        str(invoice.get("beschreibung") or ""),
    ]))
    bank_text = " ".join(filter(None, [
        str(bank.get("partner") or ""),
        str(bank.get("text") or ""),
        str(bank.get("booking_type") or ""),
    ]))
    parts.append(text_score(inv_text, bank_text))

    if invoice.get("absender") and bank.get("partner"):
        parts.append(text_score(invoice["absender"], bank["partner"]))

    return sum(parts) / len(parts) if parts else 0.0


def _candidate_payload(inv: dict, bank: dict, score: float) -> dict:
    return {
        "bank_row_id": bank["row_id"],
        "date": bank["date"].isoformat() if bank.get("date") else None,
        "amount": bank.get("amount"),
        "partner": bank.get("partner"),
        "text": bank.get("text"),
        "score": round(score, 3),
    }


def match_invoices_to_bank(
    invoices: list[dict],
    bank_rows: list[dict],
    *,
    amount_tol_abs: float = 0.02,
    amount_tol_pct: float = 0.01,
    max_days: int = 60,
    min_score: float = 0.25,
) -> dict:
    """
    invoices: [{row, beleg_nr, re_dat, absender, beschreibung, betrag}, ...]
              nur Zeilen mit leerem Zahlungsdatum
    bank_rows: normalisierte Buchungen mit selected=True

    Rückgabe:
      {
        matches, ambiguous, unmatched, no_amount,
        stats, used_bank_ids
      }
    """
    candidates_pool = [b for b in bank_rows if b.get("selected")]
    used_bank_ids = set()

    matches = []
    ambiguous = []
    unmatched = []
    no_amount = []

    for inv in invoices:
        betrag = inv.get("betrag")
        if betrag is None:
            no_amount.append({**inv, "status": STATUS_NO_AMOUNT, "reason": "kein Betrag"})
            continue

        scored = []
        for b in candidates_pool:
            if b["row_id"] in used_bank_ids:
                continue
            if not amounts_match(betrag, b.get("amount"), amount_tol_abs, amount_tol_pct):
                continue
            if not date_in_window(inv.get("re_dat"), b.get("date"), max_days):
                continue
            sc = score_candidate(inv, b)
            if sc < min_score:
                continue
            scored.append((sc, b))

        scored.sort(key=lambda x: x[0], reverse=True)

        if not scored:
            unmatched.append({**inv, "status": STATUS_NOT_FOUND})
            continue

        # Volle Härte: ≥2 treffende Kandidaten → immer mehrdeutig (kein Gap-Override)
        if len(scored) >= 2:
            ambiguous.append({
                **inv,
                "status": STATUS_AMBIGUOUS,
                "candidates": [
                    _candidate_payload(inv, b, sc) for sc, b in scored[:5]
                ],
            })
            continue

        best_score, best = scored[0]
        used_bank_ids.add(best["row_id"])
        matches.append(_result(inv, best, STATUS_FOUND, best_score))

    return {
        "matches": matches,
        "ambiguous": ambiguous,
        "unmatched": unmatched,
        "no_amount": no_amount,
        "stats": {
            "invoices": len(invoices),
            "bank_selected": len(candidates_pool),
            "gefunden": len(matches),
            "mehrdeutig": len(ambiguous),
            "nicht_gefunden": len(unmatched),
            "kein_betrag": len(no_amount),
            "bank_used": len(used_bank_ids),
        },
        "used_bank_ids": used_bank_ids,
    }


def _result(inv: dict, bank: dict, status: str, score: float) -> dict:
    text = bank.get("text") or bank.get("booking_type") or ""
    if bank.get("partner") and bank["partner"] not in text:
        display = f"{bank['partner']}: {text}" if text else bank["partner"]
    else:
        display = text or bank.get("partner") or ""
    return {
        "invoice_row": inv["row"],
        "beleg_nr": inv.get("beleg_nr"),
        "bank_row_id": bank["row_id"],
        "status": status,
        "score": round(score, 3),
        "date": bank["date"],
        "text": display[:250],
        "partner": bank.get("partner") or "",
        "amount": bank.get("amount"),
    }
