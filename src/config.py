from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv


def _app_dir() -> Path:
    """Carpeta del .exe o del proyecto: aquí van .env y reportes exportados."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _bundle_dir() -> Path:
    """Recursos empaquetados (config) dentro del ejecutable."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


ROOT_DIR = _app_dir()
CONFIG_PATH = _bundle_dir() / "config" / "field_mapping.yaml"


def load_mapping() -> dict[str, Any]:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_env() -> None:
    load_dotenv(ROOT_DIR / ".env")
    load_dotenv(ROOT_DIR / ".env.local", override=True)


def get_supabase_credentials() -> tuple[str, str]:
    load_env()
    url = os.getenv("VITE_SUPABASE_URL") or os.getenv("SUPABASE_URL")
    key = (
        os.getenv("SUPABASE_SERVICE_KEY")
        or os.getenv("SUPABASE_SERVICE_ROLE_KEY")
        or os.getenv("VITE_SUPABASE_ANON_KEY")
    )
    if not url or not key:
        env_path = ROOT_DIR / ".env"
        raise RuntimeError(
            "Faltan credenciales Supabase.\n\n"
            f"Crea el archivo:\n  {env_path}\n\n"
            "Con estas líneas:\n"
            "  VITE_SUPABASE_URL=https://tu-proyecto.supabase.co\n"
            "  SUPABASE_SERVICE_KEY=tu_clave\n\n"
            "(Copia .env.example y renómbralo a .env)"
        )
    return url, key


def default_excel_path() -> Path:
    mapping = load_mapping()
    return ROOT_DIR / mapping["excel"]["default_file"]
