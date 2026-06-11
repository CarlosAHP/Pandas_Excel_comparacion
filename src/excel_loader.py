from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from config import load_mapping


def _normalize_label(text: Any) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    return re.sub(r"\s+", " ", str(text).strip().lower())


def normalize_key(text: Any) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    return re.sub(r"\s+", "", str(text).strip().upper())


def unit_key_variants(unidad: Any) -> list[str]:
    """Variantes de unidad para cruce flexible (ej. 1502 A <-> 1502A). Sin recortar 1201 a 201."""
    if unidad is None or (isinstance(unidad, float) and pd.isna(unidad)):
        return []
    raw = str(unidad).strip().upper()
    base = normalize_key(unidad)
    if not base:
        return []
    variants = [base]
    if raw != base:
        variants.append(raw)
    return list(dict.fromkeys(variants))


def excel_join_key(proyecto: Any, unidad: Any) -> str:
    return f"{normalize_key(proyecto)}|{normalize_key(unidad)}"


def _to_float(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _read_sheet_labels(df: pd.DataFrame, label_cols: list[str]) -> dict[str, tuple[int, str]]:
    found: dict[str, tuple[int, str]] = {}
    for label_col in label_cols:
        col_idx = ord(label_col.upper()) - ord("A")
        for row_idx in range(len(df)):
            label = _normalize_label(df.iat[row_idx, col_idx] if col_idx < df.shape[1] else None)
            if label and label not in found:
                found[label] = (row_idx + 1, label_col)
    return found


def _resolve_label_row(labels: list[str], label_positions: dict[str, tuple[int, str]]) -> int | None:
    for label in labels:
        norm = _normalize_label(label)
        if norm in label_positions:
            return label_positions[norm][0]
    return None


def _find_value(
    df: pd.DataFrame,
    labels: list[str],
    value_column: str,
    label_positions: dict[str, tuple[int, str]],
) -> float | None:
    row_1 = _resolve_label_row(labels, label_positions)
    if row_1 is None:
        return None
    col_idx = ord(value_column.upper()) - ord("A")
    row_idx = row_1 - 1
    if row_idx < len(df) and col_idx < df.shape[1]:
        return _to_float(df.iat[row_idx, col_idx])
    return None


def _find_text(
    df: pd.DataFrame,
    labels: list[str],
    value_column: str,
    label_positions: dict[str, tuple[int, str]],
) -> str:
    row_1 = _resolve_label_row(labels, label_positions)
    if row_1 is None:
        return ""
    col_idx = ord(value_column.upper()) - ord("A")
    row_idx = row_1 - 1
    if row_idx < len(df) and col_idx < df.shape[1]:
        val = df.iat[row_idx, col_idx]
        if val is not None and not (isinstance(val, float) and pd.isna(val)):
            return str(val).strip()
    return ""


def _load_resumen_exclusions(path: Path, excel_cfg: dict) -> set[str]:
    """Unidades a excluir según ESTATUS en hoja RESUMEN (RE VENTA, NEGOCIACION)."""
    sheet = excel_cfg.get("resumen_sheet", "RESUMEN")
    exclude_terms = [t.upper() for t in excel_cfg.get("exclude_estatus", [])]
    excluded: set[str] = set()

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    except (ValueError, KeyError):
        return excluded

    for _, row in df.iterrows():
        estatus = str(row.iloc[0] if len(row) > 0 else "").upper()
        apto = row.iloc[1] if len(row) > 1 else None
        if not apto or (isinstance(apto, float) and pd.isna(apto)):
            continue
        if any(term in estatus for term in exclude_terms):
            for variant in unit_key_variants(apto):
                excluded.add(variant)
    return excluded


def _dedupe_by_unit(records: list[dict[str, Any]], prefer_token: str) -> list[dict[str, Any]]:
    """Si hay varias hojas por unidad, prioriza la que contiene 'correct' en el nombre."""
    by_key: dict[str, list[dict[str, Any]]] = {}
    for rec in records:
        key = rec.get("join_key") or excel_join_key(rec.get("proyecto"), rec.get("unidad"))
        by_key.setdefault(key, []).append(rec)

    result: list[dict[str, Any]] = []
    prefer = prefer_token.lower()
    for group in by_key.values():
        if len(group) == 1:
            result.append(group[0])
            continue
        with_correct = [r for r in group if prefer in str(r.get("hoja", "")).lower()]
        result.append(with_correct[0] if with_correct else group[-1])
    return result


def load_ficha_sheets(path: Path | str | None = None) -> pd.DataFrame:
    mapping = load_mapping()
    excel_cfg = mapping["excel"]
    path = Path(path) if path else Path(__file__).resolve().parent.parent / excel_cfg["default_file"]

    skip = set(excel_cfg.get("skip_sheets", []))
    pattern = excel_cfg.get("sheet_pattern", "NEO2")
    proyecto_filter = normalize_key(excel_cfg.get("proyecto_filter", "NEO2"))
    label_cols = excel_cfg.get("label_columns") or ["B"]
    fields_cfg: dict[str, Any] = excel_cfg["fields"]
    excluded_units = _load_resumen_exclusions(path, excel_cfg)

    xl = pd.ExcelFile(path)
    records: list[dict[str, Any]] = []

    for sheet_name in xl.sheet_names:
        if sheet_name in skip:
            continue
        if pattern.upper().replace(" ", "") not in sheet_name.upper().replace(" ", ""):
            continue

        df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        label_positions = _read_sheet_labels(df, label_cols)

        row: dict[str, Any] = {"hoja": sheet_name, "archivo": path.name}
        for field_name, field_def in fields_cfg.items():
            labels = field_def["labels"]
            value_col = field_def.get("value_column", excel_cfg["value_columns"]["default"])
            field_label_cols = [field_def["label_column"]] if field_def.get("label_column") else label_cols
            field_positions = {
                k: v for k, v in label_positions.items() if v[1] in field_label_cols
            } or label_positions
            if field_name in ("proyecto", "unidad"):
                row[field_name] = _find_text(df, labels, value_col, field_positions)
            else:
                row[field_name] = _find_value(df, labels, value_col, field_positions)

        if not row.get("unidad") and not row.get("proyecto"):
            continue

        if proyecto_filter and normalize_key(row.get("proyecto")) not in ("", proyecto_filter, "NEO2"):
            continue

        unidad_variants = unit_key_variants(row.get("unidad"))
        if excluded_units and any(v in excluded_units for v in unidad_variants):
            row["excluida"] = True
            continue

        row["proyecto_key"] = normalize_key(row.get("proyecto"))
        row["unidad_key"] = normalize_key(row.get("unidad"))
        row["join_key"] = excel_join_key(row.get("proyecto"), row.get("unidad"))

        # Fila 33 completa = enganche con reserva incluida (Daniel 2.1)
        row["enganche_fraccionado"] = row.get("enganche_fraccionado_mas_reserva")

        # Otros gastos Excel = solo gastoscompranofinanc (Daniel 1.4)
        row["otros_gastos"] = row.get("gastos_legales") or 0.0

        records.append(row)

    records = _dedupe_by_unit(records, excel_cfg.get("prefer_sheet_contains", "correct"))
    return pd.DataFrame(records)
