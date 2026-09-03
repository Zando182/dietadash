#!/bin/bash
# Avvia DietaDash: server locale + browser. Doppio clic da Finder.
cd "$(dirname "$0")" || exit 1

PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "Python 3 non trovato. Installalo da https://www.python.org e riprova."
  read -r -p "Premi Invio per chiudere..."
  exit 1
fi

if ! "$PY" -c "import openpyxl" >/dev/null 2>&1; then
  echo "Manca la libreria openpyxl: la installo..."
  "$PY" -m pip install --user openpyxl || {
    echo "Installazione fallita. Prova a mano:  $PY -m pip install openpyxl"
    read -r -p "Premi Invio per chiudere..."
    exit 1
  }
fi

"$PY" scripts/server.py "$1"
