"""Genera credenciales embebidas dentro del .exe (no se sube a Git)."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import dotenv_values

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "src" / "_embedded_credentials.py"


def _collect_env() -> dict[str, str]:
    merged: dict[str, str] = {}
    for name in (".env.local", ".env"):
        path = ROOT / name
        if path.exists():
            merged.update({k: v for k, v in dotenv_values(path).items() if v})

    # GitHub Actions / CI: variables de entorno
    if os.getenv("VITE_SUPABASE_URL"):
        merged["VITE_SUPABASE_URL"] = os.environ["VITE_SUPABASE_URL"]
    if os.getenv("SUPABASE_SERVICE_KEY"):
        merged["SUPABASE_SERVICE_KEY"] = os.environ["SUPABASE_SERVICE_KEY"]
    return merged


def main() -> None:
    env = _collect_env()
    url = env.get("VITE_SUPABASE_URL") or env.get("SUPABASE_URL")
    key = env.get("SUPABASE_SERVICE_KEY") or env.get("SUPABASE_SERVICE_ROLE_KEY")

    if not url or not key:
        raise SystemExit(
            "ERROR: faltan credenciales para empaquetar el .exe.\n"
            "Crea .env o .env.local con VITE_SUPABASE_URL y SUPABASE_SERVICE_KEY,\n"
            "o define esas variables en GitHub Secrets (Actions)."
        )

    OUT.write_text(
        f'''# AUTO-GENERADO por prepare_build.py — NO SUBIR A GIT
SUPABASE_URL = {url!r}
SUPABASE_KEY = {key!r}
''',
        encoding="utf-8",
    )
    print(f"OK: credenciales listas para empaquetar en {OUT.name}")


if __name__ == "__main__":
    main()
