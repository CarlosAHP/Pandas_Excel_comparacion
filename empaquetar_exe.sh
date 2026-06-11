#!/bin/bash
# Desde Mac: prepara credenciales y dispara build en GitHub (genera .exe con todo dentro)
set -e
cd "$(dirname "$0")"

if [ ! -f .env ] && [ ! -f .env.local ]; then
  echo "ERROR: necesitas .env o .env.local con credenciales Supabase"
  exit 1
fi

if ! command -v gh >/dev/null 2>&1; then
  echo "Instala GitHub CLI: brew install gh && gh auth login"
  exit 1
fi

echo "=== Configurando secrets en GitHub (una sola vez) ==="
URL=$(grep -E '^VITE_SUPABASE_URL=' .env.local .env 2>/dev/null | head -1 | cut -d= -f2-)
KEY=$(grep -E '^SUPABASE_SERVICE_KEY=' .env.local .env 2>/dev/null | head -1 | cut -d= -f2-)

if [ -z "$URL" ] || [ -z "$KEY" ]; then
  echo "ERROR: .env debe tener VITE_SUPABASE_URL y SUPABASE_SERVICE_KEY"
  exit 1
fi

gh secret set VITE_SUPABASE_URL --body "$URL"
gh secret set SUPABASE_SERVICE_KEY --body "$KEY"
echo "Secrets configurados."

echo ""
echo "=== Iniciando build del .exe en GitHub ==="
gh workflow run build-windows.yml
echo ""
echo "Ve a: https://github.com/CarlosAHP/Pandas_Excel_comparacion/actions"
echo "Cuando termine (~5 min), descarga ComparadorCotizaciones-Windows"
echo "Sube SOLO ComparadorCotizaciones.exe a Google Drive."
