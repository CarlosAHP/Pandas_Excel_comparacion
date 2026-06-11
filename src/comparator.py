from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from config import load_mapping
from excel_loader import excel_join_key, normalize_key
from supabase_client import lookup_cotizacion


@dataclass
class ComparisonRow:
    join_key: str
    proyecto: str
    unidad: str
    nivel: str
    campo: str
    valor_excel: float | None
    valor_supabase: float | None
    diferencia: float | None
    estado: str
    id_crmtratocotizacion: str | None = None
    nota: str = ""
    seccion: str = "desglose"  # desglose | validacion | referencia


@dataclass
class UnitSummary:
    join_key: str
    proyecto: str
    unidad: str
    total_campos: int = 0
    ok: int = 0
    diferencias: int = 0
    sin_dato: int = 0
    estado_unidad: str = "OK"
    id_crmtratocotizacion: str | None = None
    total_excel: float | None = None
    total_supabase: float | None = None

    @property
    def tiene_error(self) -> bool:
        return self.diferencias > 0 or self.sin_dato > 0


@dataclass
class ComparisonResult:
    rows: list[ComparisonRow] = field(default_factory=list)
    units: list[UnitSummary] = field(default_factory=list)
    excel_sin_match: list[str] = field(default_factory=list)
    excluidas: list[str] = field(default_factory=list)

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([r.__dict__ for r in self.rows])

    def units_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([u.__dict__ for u in self.units])

    @property
    def resumen(self) -> dict[str, int]:
        df = self.to_dataframe()
        if df.empty:
            return {"unidades": 0, "total": 0, "ok": 0, "diferencia": 0, "sin_dato": 0}
        unidades_ok = sum(1 for u in self.units if u.estado_unidad == "OK")
        unidades_error = sum(1 for u in self.units if u.estado_unidad == "DIFERENCIA")
        return {
            "unidades": len(self.units),
            "unidades_ok": unidades_ok,
            "unidades_con_diferencia": unidades_error,
            "total": len(df),
            "ok": int((df["estado"] == "OK").sum()),
            "diferencia": int((df["estado"] == "DIFERENCIA").sum()),
            "sin_dato": int((df["estado"] == "SIN_DATO").sum()),
        }


def _f(val: Any) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _compare_values(
    excel_val: float | None,
    supa_val: float | None,
    tolerance: float,
) -> tuple[float | None, str]:
    if excel_val is None and supa_val is None:
        return None, "SIN_DATO"
    if excel_val is None or supa_val is None:
        return None, "SIN_DATO"
    diff = round(excel_val - supa_val, 2)
    if abs(diff) <= tolerance:
        return diff, "OK"
    return diff, "DIFERENCIA"


def _supabase_enganche(row: pd.Series) -> float | None:
    """Enganche Supabase = engancheafraccionar + montoreserva (equivale a fila 33 Excel)."""
    eng = _f(row.get("engancheafraccionar_crmtratocotizacion"))
    res = _f(row.get("montoreserva_crmtratocotizacion"))
    if eng is None and res is None:
        return None
    return round((eng or 0) + (res or 0), 2)


def _supabase_otros_gastos(row: pd.Series) -> float | None:
    g = _f(row.get("gastoscompranofinanc_crmtratocotizacion"))
    o = _f(row.get("otroscargosaddnofinanc_crmtratocotizacion"))
    if g is None and o is None:
        return None
    total = (g or 0) + (o or 0)
    return total if total else None


def _cotizacion_values(row: pd.Series) -> dict[str, float | None]:
    return {
        "reserva": _f(row.get("montoreserva_crmtratocotizacion")),
        "abono_pvc": _f(row.get("montopromesa_crmtratocotizacion")),
        "enganche_fraccionado": _supabase_enganche(row),
        "monto_financiar": _f(row.get("montofinanciar_crmtratocotizacion")),
        "otros_gastos": _supabase_otros_gastos(row),
        "precio_total_con_gastos": _f(row.get("preciototalygastoscompranofin_crmtratocotizacion")),
        "precio_venta": _f(row.get("preciototalventa_crmtratocotizacion")),
    }


def _excel_values(row: pd.Series) -> dict[str, float | None]:
    return {
        "reserva": _f(row.get("reserva")),
        "abono_pvc": _f(row.get("abono_pvc")),
        "enganche_fraccionado": _f(row.get("enganche_fraccionado")),
        "monto_financiar": _f(row.get("monto_financiar")),
        "otros_gastos": _f(row.get("otros_gastos")),
        "precio_total_con_gastos": _f(row.get("precio_total_con_gastos")),
        "precio_venta": _f(row.get("precio_venta")),
    }


def _suma_desglose_excel(vals: dict[str, float | None]) -> float | None:
    """Reserva + PVC + enganche (sin doble contar reserva en fila 33) + financiar + otros."""
    reserva = vals.get("reserva")
    pvc = vals.get("abono_pvc")
    eng_fila33 = vals.get("enganche_fraccionado")
    fin = vals.get("monto_financiar")
    otros = vals.get("otros_gastos")
    if any(v is None for v in (reserva, pvc, eng_fila33, fin, otros)):
        return None
    eng_sin_reserva = max(0.0, eng_fila33 - reserva)  # type: ignore[operator]
    return round(reserva + pvc + eng_sin_reserva + fin + otros, 2)  # type: ignore[operator]


def _suma_desglose_supabase(row: pd.Series) -> float | None:
    """Fórmula Daniel: reserva + pvc + engancheafraccionar + financiar + otros_gastos."""
    reserva = _f(row.get("montoreserva_crmtratocotizacion"))
    pvc = _f(row.get("montopromesa_crmtratocotizacion"))
    eng = _f(row.get("engancheafraccionar_crmtratocotizacion"))
    fin = _f(row.get("montofinanciar_crmtratocotizacion"))
    otros = _supabase_otros_gastos(row)
    if any(v is None for v in (reserva, pvc, eng, fin, otros)):
        return None
    return round(reserva + pvc + eng + fin + otros, 2)  # type: ignore[operator]


def compare_all(excel_df: pd.DataFrame, cotizaciones_df: pd.DataFrame) -> ComparisonResult:
    mapping = load_mapping()
    tolerance = float(mapping["comparison"]["tolerance"])
    result = ComparisonResult()

    if excel_df.empty:
        return result

    compare_fields = list(mapping["comparison"]["concept_keys"])
    if "suma_desglose" in compare_fields:
        compare_fields = [f for f in compare_fields if f != "suma_desglose"]

    unit_rows: dict[str, list[ComparisonRow]] = {}

    for _, ex_row in excel_df.iterrows():
        jk = ex_row.get("join_key") or excel_join_key(ex_row.get("proyecto"), ex_row.get("unidad"))
        proyecto = str(ex_row.get("proyecto", ""))
        unidad = str(ex_row.get("unidad", ""))
        ex_vals = _excel_values(ex_row)

        cot_row = lookup_cotizacion(
            cotizaciones_df,
            normalize_key(ex_row.get("proyecto_key") or "NEO2"),
            ex_row.get("unidad"),
        )

        if cot_row is None:
            result.excel_sin_match.append(f"{proyecto} / {unidad} ({jk})")
            continue

        cot_id = str(cot_row["id_crmtratocotizacion"])
        cot_vals = _cotizacion_values(cot_row)

        for field_name in compare_fields:
            excel_val = ex_vals.get(field_name)
            supa_val = cot_vals.get(field_name)
            diff, estado = _compare_values(excel_val, supa_val, tolerance)
            seccion = "desglose" if field_name != "precio_venta" else "referencia"
            row = ComparisonRow(
                join_key=jk,
                proyecto=proyecto,
                unidad=unidad,
                nivel="Excel vs Cotización",
                campo=field_name,
                valor_excel=excel_val,
                valor_supabase=supa_val,
                diferencia=diff,
                estado=estado,
                id_crmtratocotizacion=cot_id,
                seccion=seccion,
            )
            result.rows.append(row)
            unit_rows.setdefault(jk, []).append(row)

        # Validación suma desglose vs total (Excel vs Supabase por separado, informativo)
        suma_ex = _suma_desglose_excel(ex_vals)
        total_ex = ex_vals.get("precio_total_con_gastos")
        diff_s, est_s = _compare_values(suma_ex, total_ex, tolerance)
        row_s = ComparisonRow(
            join_key=jk,
            proyecto=proyecto,
            unidad=unidad,
            nivel="Excel vs Cotización",
            campo="suma_desglose",
            valor_excel=suma_ex,
            valor_supabase=total_ex,
            diferencia=diff_s,
            estado=est_s,
            id_crmtratocotizacion=cot_id,
            nota="Suma conceptos Excel vs Precio total Excel",
            seccion="validacion",
        )
        result.rows.append(row_s)
        unit_rows.setdefault(jk, []).append(row_s)

        suma_supa = _suma_desglose_supabase(cot_row)
        total_supa = cot_vals.get("precio_total_con_gastos")
        diff_c, est_c = _compare_values(suma_supa, total_supa, tolerance)
        row_c = ComparisonRow(
            join_key=jk,
            proyecto=proyecto,
            unidad=unidad,
            nivel="Excel vs Cotización",
            campo="suma_desglose_supabase",
            valor_excel=suma_supa,
            valor_supabase=total_supa,
            diferencia=diff_c,
            estado=est_c,
            id_crmtratocotizacion=cot_id,
            nota="Suma conceptos Supabase vs preciototalygastoscompranofin",
            seccion="validacion",
        )
        result.rows.append(row_c)
        unit_rows.setdefault(jk, []).append(row_c)

    main_fields = {
        "reserva",
        "abono_pvc",
        "enganche_fraccionado",
        "monto_financiar",
        "otros_gastos",
        "precio_total_con_gastos",
    }

    for jk, rows in unit_rows.items():
        first = rows[0]
        main = [r for r in rows if r.campo in main_fields]
        ok = sum(1 for r in main if r.estado == "OK")
        diff = sum(1 for r in main if r.estado == "DIFERENCIA")
        sin = sum(1 for r in main if r.estado == "SIN_DATO")
        total_row = next((r for r in main if r.campo == "precio_total_con_gastos"), None)
        result.units.append(
            UnitSummary(
                join_key=jk,
                proyecto=first.proyecto,
                unidad=first.unidad,
                total_campos=len(main),
                ok=ok,
                diferencias=diff,
                sin_dato=sin,
                estado_unidad="OK" if diff == 0 and sin == 0 else "DIFERENCIA",
                id_crmtratocotizacion=first.id_crmtratocotizacion,
                total_excel=total_row.valor_excel if total_row else None,
                total_supabase=total_row.valor_supabase if total_row else None,
            )
        )

    result.units.sort(key=lambda u: (u.estado_unidad != "DIFERENCIA", u.unidad))
    return result
