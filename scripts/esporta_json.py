#!/usr/bin/env python3
"""
Scrive web/dati.json: l'istantanea del workbook che fa funzionare la
dashboard anche senza server, per esempio pubblicata su GitHub Pages.

Online la pagina e' in sola lettura: il salvataggio ha bisogno del server
locale, che sul sito non c'e'. Serve a guardare i numeri da telefono o a far
vedere la dashboard a qualcuno.

    python3 scripts/esporta_json.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import dieta_excel as db  # noqa: E402

DESTINAZIONE = Path(__file__).resolve().parent.parent / "web" / "dati.json"


def main() -> int:
    try:
        dati = db.leggi()
    except db.ErroreDati as e:
        print(f"Non riesco a leggere il workbook: {e}", file=sys.stderr)
        return 1

    # Il percorso del file sul disco non serve online e direbbe a chiunque
    # come si chiama il tuo utente: fuori.
    dati.pop("file", None)
    # Le misurazioni (peso, circonferenze) restano sul computer: la dashboard
    # non le mostra da nessuna parte, quindi pubblicarle non servirebbe a nulla.
    misure = len(dati.pop("misure", []))
    dati["solaLettura"] = True

    DESTINAZIONE.write_text(json.dumps(dati, ensure_ascii=False), encoding="utf-8")
    peso = DESTINAZIONE.stat().st_size / 1024
    print(f"Scritto {DESTINAZIONE.name}: {len(dati['giorni'])} giorni, {peso:.0f} kB")
    print(f"Escluse {misure} misurazioni: quelle non escono dal tuo computer.")
    print("Ricordati che questo file finisce online: contiene tutto il registro alimentare.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
