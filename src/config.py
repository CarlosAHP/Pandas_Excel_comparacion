from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

ROOT_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT_DIR / "config" / "field_mapping.yaml"


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
        raise RuntimeError(
            "Faltan credenciales Supabase. Define VITE_SUPABASE_URL y SUPABASE_SERVICE_KEY en .env/.env.local"
        )
    return url, key


def default_excel_path() -> Path:
    mapping = load_mapping()
    return ROOT_DIR / mapping["excel"]["default_file"]
