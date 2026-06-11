#!/bin/bash
# Arranque con el Python del venv (debe ser 3.12+ con Tk 8.6+)
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  echo "Creando entorno virtual con Python 3.12..."
  python3.12 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi
exec .venv/bin/python src/main.py
