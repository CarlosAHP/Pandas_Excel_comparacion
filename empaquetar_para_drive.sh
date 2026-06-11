#!/bin/bash
# Crea un ZIP liviano para subir a Google Drive (sin paquetes ni credenciales)
set -e
cd "$(dirname "$0")"
NOMBRE="Pandas_Excel_comparacion"
ARCHIVO="../${NOMBRE}.zip"

rm -f "$ARCHIVO"
zip -r "$ARCHIVO" . \
  -x ".venv/*" \
  -x ".git/*" \
  -x ".env" \
  -x ".env.local" \
  -x "*.xlsx" \
  -x ".DS_Store" \
  -x "__pycache__/*" \
  -x "*/__pycache__/*" \
  -x "reporte_diferencias*.xlsx"

echo ""
echo "Listo: $ARCHIVO"
echo "Sube ese ZIP a Google Drive. El equipo lo descarga y descomprime."
echo "NO incluye .env — cada quien crea el suyo con las credenciales."
