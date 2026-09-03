#!/usr/bin/env python3
"""
Server locale della dashboard Dieta.

Serve la pagina in web/ e espone il workbook come database:

    GET  /api/dati        registro completo + misure + impostazioni
    GET  /api/revisione   solo il timbro di modifica del file (polling leggero)
    POST /api/giorni      inserisce o aggiorna uno o piu' giorni
    POST /api/config      salva le impostazioni (es. data del prossimo check)
    DELETE /api/giorno?data=AAAA-MM-GG

Il doppio mandato sta qui: la pagina scrive chiamando queste rotte, e si
accorge delle modifiche fatte a mano in Excel confrontando /api/revisione.

    python3 scripts/server.py [porta]
"""

from __future__ import annotations

import json
import mimetypes
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))

import dieta_excel as db  # noqa: E402

RADICE = Path(__file__).resolve().parent.parent
WEB = RADICE / "web"
PORTA = int(sys.argv[1]) if len(sys.argv) > 1 else 5191
CORPO_MASSIMO = 4 * 1024 * 1024

# Un solo scrittore per volta: due salvataggi simultanei riscriverebbero
# entrambi l'intero foglio partendo dalla stessa lettura, e uno andrebbe perso.
LUCCHETTO = threading.Lock()


class Gestore(BaseHTTPRequestHandler):
    server_version = "DietaDash"

    def log_message(self, formato, *args):  # meno rumore in console
        if "/api/revisione" not in (self.path or ""):
            sys.stderr.write(f"  {self.address_string()} {formato % args}\n")

    # ---------------------------------------------------------------- utili

    def _json(self, stato: int, corpo: dict) -> None:
        testo = json.dumps(corpo, ensure_ascii=False).encode("utf-8")
        self.send_response(stato)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(testo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(testo)

    def _corpo(self) -> dict:
        lunghezza = int(self.headers.get("Content-Length") or 0)
        if lunghezza > CORPO_MASSIMO:
            raise db.ErroreDati("Richiesta troppo grande.")
        if lunghezza == 0:
            return {}
        try:
            return json.loads(self.rfile.read(lunghezza).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as e:
            raise db.ErroreDati("Corpo della richiesta non leggibile.") from e

    def _statico(self, percorso: str) -> None:
        rel = percorso.lstrip("/") or "index.html"
        file = (WEB / rel).resolve()
        # resolve() + prefisso: impedisce di uscire da web/ con ../
        if not str(file).startswith(str(WEB)) or not file.is_file():
            file = WEB / "index.html"
        if not file.is_file():
            self._json(404, {"errore": "Pagina non trovata: manca web/index.html"})
            return
        dati = file.read_bytes()
        tipo = mimetypes.guess_type(file.name)[0] or "application/octet-stream"
        if tipo.startswith("text/") or tipo.endswith(("javascript", "json")):
            tipo += "; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(dati)))
        # no-store e non no-cache: senza un validatore (ETag o Last-Modified)
        # il browser puo' comunque riusare la copia in memoria, e dopo una
        # modifica a web/ si continuerebbe a vedere la versione vecchia.
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.end_headers()
        self.wfile.write(dati)

    def _gestisci(self, azione) -> None:
        try:
            stato, corpo = azione()
        except db.ErroreDati as e:
            self._json(409, {"errore": str(e)})
        except Exception as e:  # imprevisto: la pagina mostra il messaggio
            self._json(500, {"errore": f"{type(e).__name__}: {e}"})
        else:
            self._json(stato, corpo)

    # --------------------------------------------------------------- rotte

    def do_GET(self) -> None:
        percorso = urlparse(self.path).path
        if percorso == "/api/dati":
            self._gestisci(lambda: (200, db.leggi()))
        elif percorso == "/api/revisione":
            self._gestisci(lambda: (200, db.revisione()))
        elif percorso == "/api/stato":
            self._gestisci(lambda: (200, db.stato()))
        elif percorso.startswith("/api/"):
            self._json(404, {"errore": f"Rotta sconosciuta: GET {percorso}"})
        else:
            self._statico(percorso)

    def do_POST(self) -> None:
        percorso = urlparse(self.path).path

        if percorso == "/api/giorni":
            def azione():
                corpo = self._corpo()
                giorni = corpo.get("giorni")
                if giorni is None and corpo.get("data"):
                    giorni = [corpo]  # comodita': un giorno solo, senza involucro
                if not isinstance(giorni, list) or not giorni:
                    raise db.ErroreDati('Serve un elenco "giorni" non vuoto.')
                with LUCCHETTO:
                    esito = db.salva_giorni(giorni)
                    esito["dati"] = db.leggi()
                return 200, esito

            self._gestisci(azione)

        elif percorso == "/api/config":
            def azione():
                corpo = self._corpo()
                voci = corpo.get("config", corpo)
                if not isinstance(voci, dict) or not voci:
                    raise db.ErroreDati("Nessuna impostazione da salvare.")
                with LUCCHETTO:
                    return 200, db.salva_config({str(k): v for k, v in voci.items()})

            self._gestisci(azione)

        else:
            self._json(404, {"errore": f"Rotta sconosciuta: POST {percorso}"})

    def do_DELETE(self) -> None:
        parti = urlparse(self.path)
        if parti.path != "/api/giorno":
            self._json(404, {"errore": f"Rotta sconosciuta: DELETE {parti.path}"})
            return

        def azione():
            data = (parse_qs(parti.query).get("data") or [""])[0]
            if not data:
                raise db.ErroreDati("Serve il parametro ?data=AAAA-MM-GG.")
            with LUCCHETTO:
                esito = db.elimina_giorno(data)
                esito["dati"] = db.leggi()
            return 200, esito

        self._gestisci(azione)


def main() -> int:
    # Il workbook non sta nel repository (contiene dati personali): al primo
    # avvio se ne crea uno vuoto, cosi' chi scarica il progetto parte subito.
    if db.crea_se_manca():
        print(f"Non c'era nessun workbook: ne ho creato uno vuoto in {db.WORKBOOK.name}.")

    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", PORTA), Gestore)
    except OSError as e:
        if getattr(e, "errno", None) in (48, 98):
            print(f"Porta {PORTA} occupata. Riprova con: python3 scripts/server.py {PORTA + 1}")
            return 1
        raise

    indirizzo = f"http://localhost:{PORTA}"
    print(f"DietaDash su {indirizzo}   (Ctrl+C per chiudere)")
    print(f"Database: {db.WORKBOOK}")
    print("Quello che inserisci nella pagina finisce nell'Excel; se modifichi")
    print("l'Excel a mano, la pagina se ne accorge da sola entro pochi secondi.")
    threading.Timer(0.6, lambda: webbrowser.open(indirizzo)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nChiuso.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
