from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

try:
    from _embedded_credentials import SUPABASE_KEY as _EMBEDDED_KEY
    from _embedded_credentials import SUPABASE_URL as _EMBEDDED_URL
except ImportError:
    _EMBEDDED_URL = None
    _EMBEDDED_KEY = None


def _app_dir() -> Path:
    """Carpeta del .exe o del proyecto: reportes exportados."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _bundle_dir() -> Path:
    """Recursos empaquetados dentro del ejecutable."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


ROOT_DIR = _app_dir()
CONFIG_PATH = _bundle_dir() / "config" / "field_mapping.yaml"


def load_mapping() -> dict[str, Any]:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def _apply_embedded_credentials() -> bool:
    if _EMBEDDED_URL and _EMBEDDED_KEY:
        os.environ.setdefault("VITE_SUPABASE_URL", _EMBEDDED_URL)
        os.environ.setdefault("SUPABASE_SERVICE_KEY", _EMBEDDED_KEY)
        return True
    return False


def load_env() -> None:
    if getattr(sys, "frozen", False) and _apply_embedded_credentials():
        return
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
        if getattr(sys, "frozen", False):
            raise RuntimeError(
                "Este ejecutable se generó sin credenciales embebidas.\n"
                "Pide al equipo de desarrollo un .exe nuevo."
            )
        env_path = ROOT_DIR / ".env"
        raise RuntimeError(
            "Faltan credenciales Supabase.\n\n"
            f"Crea el archivo:\n  {env_path}\n\n"
            "Con estas líneas:\n"
            "  VITE_SUPABASE_URL=https://tu-proyecto.supabase.co\n"
            "  SUPABASE_SERVICE_KEY=tu_clave"
        )
    return url, key


def default_excel_path() -> Path:
    mapping = load_mapping()
    return ROOT_DIR / mapping["excel"]["default_file"]
