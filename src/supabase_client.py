from __future__ import annotations

from typing import Any

import pandas as pd
from supabase import Client, create_client

from config import get_supabase_credentials, load_mapping
from excel_loader import normalize_key, unit_key_variants

COTIZACION_SELECT = (
    "id_crmtratocotizacion,createdat_crmtratocotizacion,"
    "nombreproyecto_producto,nombre_producto,lote_producto,"
    "montoreserva_crmtratocotizacion,montopromesa_crmtratocotizacion,"
    "engancheafraccionar_crmtratocotizacion,montofinanciar_crmtratocotizacion,"
    "gastoscompranofinanc_crmtratocotizacion,otroscargosaddnofinanc_crmtratocotizacion,"
    "preciototalygastoscompranofin_crmtratocotizacion,preciototalventa_crmtratocotizacion"
)

PAGE_SIZE = 1000


def _client() -> Client:
    url, key = get_supabase_credentials()
    return create_client(url, key)


def _paginate_table(client: Client, table: str, select: str) -> list[dict]:
    rows: list[dict] = []
    offset = 0
    while True:
        response = client.table(table).select(select).range(offset, offset + PAGE_SIZE - 1).execute()
        batch = response.data or []
        if not batch:
            break
        rows.extend(batch)
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return rows


def _paginate_in(client: Client, table: str, select: str, column: str, ids: list[str]) -> list[dict]:
    if not ids:
        return []
    rows: list[dict] = []
    for i in range(0, len(ids), 100):
        chunk = ids[i : i + 100]
        response = client.table(table).select(select).in_(column, chunk).execute()
        rows.extend(response.data or [])
    return rows


def _join_keys_for_row(row: pd.Series) -> list[str]:
    proyecto = normalize_key(row.get("nombreproyecto_producto"))
    keys: list[str] = []
    for col in ("nombre_producto", "lote_producto"):
        val = row.get(col)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        for variant in unit_key_variants(val):
            keys.append(f"{proyecto}|{variant}")
    return list(dict.fromkeys(keys))


def _reserva_priority(estatus: str | None) -> int:
    """Menor = mejor. Prioriza reserva Autorizada (Daniel 3.2)."""
    e = (estatus or "").upper()
    if "AUTORIZAD" in e:
        return 0
    if "PROMESA" in e or "FIRMAD" in e:
        return 1
    if "PENDIENTE" in e:
        return 2
    return 3


def fetch_reservas_activas() -> pd.DataFrame:
    client = _client()
    reservas = _paginate_table(
        client,
        "crm_tratos_reservas",
        "crmtratoreserva_crmtratocotizacion_id,isactive_crmtratoreserva,estatus_crmtratoreserva",
    )
    rows = [r for r in reservas if r.get("isactive_crmtratoreserva") is True]
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def fetch_cotizaciones() -> pd.DataFrame:
    mapping = load_mapping()
    proyecto_filter = normalize_key(mapping["supabase"].get("proyecto_filter", "NEO2"))

    client = _client()
    raw = _paginate_table(client, "vista_crm_tratos_cotizaciones", COTIZACION_SELECT)
    if not raw:
        return pd.DataFrame()

    df = pd.DataFrame(raw)
    df["proyecto_key"] = df["nombreproyecto_producto"].map(normalize_key)
    df = df[df["proyecto_key"] == proyecto_filter].copy()

    reservas_df = fetch_reservas_activas()
    if reservas_df.empty:
        return pd.DataFrame()

    active_ids = set(reservas_df["crmtratoreserva_crmtratocotizacion_id"].astype(str))
    estatus_map = reservas_df.set_index("crmtratoreserva_crmtratocotizacion_id")[
        "estatus_crmtratoreserva"
    ].to_dict()

    df = df[df["id_crmtratocotizacion"].astype(str).isin(active_ids)].copy()
    df["reserva_estatus"] = df["id_crmtratocotizacion"].map(estatus_map)
    df["reserva_priority"] = df["reserva_estatus"].map(_reserva_priority)

    df["createdat_crmtratocotizacion"] = pd.to_datetime(
        df["createdat_crmtratocotizacion"], errors="coerce"
    )

    # Una cotización por clave; prioriza Autorizada, luego más reciente
    picked: list[pd.Series] = []
    key_to_rows: dict[str, list[pd.Series]] = {}
    for _, row in df.iterrows():
        for jk in _join_keys_for_row(row):
            key_to_rows.setdefault(jk, []).append(row)

    for rows in key_to_rows.values():
        best = sorted(
            rows,
            key=lambda r: (r["reserva_priority"], -r["createdat_crmtratocotizacion"].value),
        )[0]
        picked.append(best)

    if not picked:
        return pd.DataFrame()

    result = pd.DataFrame(picked).drop_duplicates("id_crmtratocotizacion")
    result["join_keys"] = result.apply(_join_keys_for_row, axis=1)
    return result.reset_index(drop=True)


def lookup_cotizacion(cotizaciones_df: pd.DataFrame, proyecto_key: str, unidad: Any) -> pd.Series | None:
    """Busca cotización por proyecto + unidad; prioriza coincidencia exacta de clave."""
    if cotizaciones_df.empty:
        return None
    proyecto_key = normalize_key(proyecto_key) or normalize_key("NEO2")
    variants = unit_key_variants(unidad)
    candidates = [f"{proyecto_key}|{v}" for v in variants]
    primary = candidates[0] if candidates else None

    exact_match: pd.Series | None = None
    fuzzy_match: pd.Series | None = None

    for _, row in cotizaciones_df.iterrows():
        row_keys = row.get("join_keys") or _join_keys_for_row(row)
        if primary and primary in row_keys:
            exact_match = row
            break
        if any(c in row_keys for c in candidates):
            fuzzy_match = row

    return exact_match if exact_match is not None else fuzzy_match


def load_supabase_bundle() -> pd.DataFrame:
    """Solo cotizaciones NEO2 con reserva activa (sin cadena plan/cuotas en v1)."""
    return fetch_cotizaciones()
