"""
Matching Kontoauszug-Buchungen ↔ Rechnungszeilen (Issue #25).

Regeln:
  1. Betrag (± Toleranz)
  2. Datum-Fenster (Rechnungsdatum … + max_days)
  3. Fuzzy-Text (Absender / Beschreibung ↔ Partner / Buchungstext)

Nur Rechnungszeilen mit leerem Zahlungsdatum (Spalte C).
"""

from __future__ import annotations

import re
from datetime import date, timedelta
from difflib import SequenceMatcher
from typing import Optional


STATUS_FOUND = "gefunden"
STATUS_AMBIGUOUS = "mehrdeutig"
STATUS_NOT_FOUND = "nicht gefunden"


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
    # Token-Overlap
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

    # Partner vs Absender extra gewichten
    if invoice.get("absender") and bank.get("partner"):
        parts.append(text_score(invoice["absender"], bank["partner"]))

    return sum(parts) / len(parts) if parts else 0.0


def match_invoices_to_bank(
    invoices: list[dict],
    bank_rows: list[dict],
    *,
    amount_tol_abs: float = 0.02,
    amount_tol_pct: float = 0.01,
    max_days: int = 60,
    min_score: float = 0.25,
    ambiguous_gap: float = 0.08,
) -> dict:
    """
    invoices: [{row, beleg_nr, re_dat, absender, beschreibung, betrag}, ...]
              nur Zeilen mit leerem Zahlungsdatum
    bank_rows: normalisierte Buchungen mit selected=True

    Rückgabe:
      {
        matches: [{invoice_row, bank_row_id, status, score, date, text, partner, amount}],
        ambiguous: [...],
        unmatched: [{invoice_row, beleg_nr, ...}],
        stats: {...}
      }
    """
    candidates_pool = [b for b in bank_rows if b.get("selected")]
    used_bank_ids = set()

    matches = []
    ambiguous = []
    unmatched = []

    for inv in invoices:
        betrag = inv.get("betrag")
        if betrag is None:
            unmatched.append({**inv, "status": STATUS_NOT_FOUND, "reason": "kein Betrag"})
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
            scored.append((sc, b))

        scored.sort(key=lambda x: x[0], reverse=True)

        if not scored or scored[0][0] < min_score:
            # Betrags-Treffer ohne Text? Wenn genau ein Betrags+Datum-Treffer → trotzdem gefunden
            amount_only = []
            for b in candidates_pool:
                if b["row_id"] in used_bank_ids:
                    continue
                if not amounts_match(betrag, b.get("amount"), amount_tol_abs, amount_tol_pct):
                    continue
                if not date_in_window(inv.get("re_dat"), b.get("date"), max_days):
                    continue
                amount_only.append(b)
            if len(amount_only) == 1:
                b = amount_only[0]
                used_bank_ids.add(b["row_id"])
                matches.append(_result(inv, b, STATUS_FOUND, 0.5))
            elif len(amount_only) > 1:
                ambiguous.append({
                    **inv,
                    "status": STATUS_AMBIGUOUS,
                    "candidates": [
                        {
                            "bank_row_id": x["row_id"],
                            "date": x["date"].isoformat() if x.get("date") else None,
                            "amount": x.get("amount"),
                            "partner": x.get("partner"),
                            "text": x.get("text"),
                            "score": round(score_candidate(inv, x), 3),
                        }
                        for x in amount_only[:5]
                    ],
                })
            else:
                unmatched.append({**inv, "status": STATUS_NOT_FOUND})
            continue

        best_score, best = scored[0]
        second = scored[1][0] if len(scored) > 1 else 0.0

        if len(scored) > 1 and (best_score - second) < ambiguous_gap and second >= min_score:
            ambiguous.append({
                **inv,
                "status": STATUS_AMBIGUOUS,
                "candidates": [
                    {
                        "bank_row_id": b["row_id"],
                        "date": b["date"].isoformat() if b.get("date") else None,
                        "amount": b.get("amount"),
                        "partner": b.get("partner"),
                        "text": b.get("text"),
                        "score": round(sc, 3),
                    }
                    for sc, b in scored[:5]
                ],
            })
        else:
            used_bank_ids.add(best["row_id"])
            matches.append(_result(inv, best, STATUS_FOUND, best_score))

    return {
        "matches": matches,
        "ambiguous": ambiguous,
        "unmatched": unmatched,
        "stats": {
            "invoices": len(invoices),
            "bank_selected": len(candidates_pool),
            "gefunden": len(matches),
            "mehrdeutig": len(ambiguous),
            "nicht_gefunden": len(unmatched),
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
