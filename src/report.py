from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from comparator import ComparisonResult, ComparisonRow
from config import load_mapping

# --- Estilos ---
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
UNIT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

RED_FONT = Font(color="9C0006", bold=True)
GREEN_FONT = Font(color="1A7F37", bold=True)
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, color="1F3864")
UNIT_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14)
SUBTITLE_FONT = Font(bold=True, size=11, color="444444")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY_FMT = '#,##0.00'

DESGLOSE_LABELS = {
    "reserva": "1. Reserva",
    "abono_pvc": "2. Abono PVC (promesa de compraventa)",
    "enganche_fraccionado": "3. Enganche fraccionado",
    "monto_financiar": "4. Monto a financiar",
    "otros_gastos": "5. Otros gastos",
    "precio_total_con_gastos": "6. Precio total con gastos",
    "precio_venta": "Precio de venta (referencia)",
    "suma_desglose": "Validación Excel: suma conceptos vs total",
    "suma_desglose_supabase": "Validación Supabase: suma conceptos vs total",
}

ACCION_POR_CAMPO = {
    "reserva": "Revisar montoreserva_crmtratocotizacion en la cotización vs fila Reserva del Excel.",
    "abono_pvc": "Revisar montopromesa_crmtratocotizacion vs 'Al firmar promesa de compraventa' en Excel.",
    "enganche_fraccionado": (
        "En Supabase: engancheafraccionar + montoreserva. "
        "En Excel: fila 33 'Enganche fraccionado mas reserva'."
    ),
    "monto_financiar": "Revisar montofinanciar_crmtratocotizacion vs 'Monto a financiar' (col. H/I) en Excel.",
    "otros_gastos": (
        "En Supabase: gastoscompranofinanc + otroscargosaddnofinanc. "
        "En Excel solo 'Gastos y honorarios legales' (gastoscompranofinanc)."
    ),
    "precio_total_con_gastos": (
        "Revisar preciototalygastoscompranofin_crmtratocotizacion vs 'Precio total' en Excel."
    ),
    "precio_venta": "Referencia: preciototalventa_crmtratocotizacion vs 'Precio de venta' en Excel.",
    "suma_desglose": "La suma de conceptos en Excel no cuadra con el Precio total de la ficha.",
    "suma_desglose_supabase": "La suma de conceptos en Supabase no cuadra con preciototalygastoscompranofin.",
}


def _mapping_reference() -> pd.DataFrame:
    mapping = load_mapping()
    sb = mapping["supabase"]["cotizacion_fields"]
    rows = [
        {
            "concepto": "Reserva",
            "excel_etiqueta": "Reserva",
            "excel_columna": "C",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": sb["reserva"],
            "formula_supabase": "Valor directo",
            "notas": "Se cuenta una sola vez en la suma total.",
        },
        {
            "concepto": "Abono PVC",
            "excel_etiqueta": "Al firmar promesa de compraventa",
            "excel_columna": "C",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": sb["abono_pvc"],
            "formula_supabase": "Valor directo",
            "notas": "",
        },
        {
            "concepto": "Enganche fraccionado",
            "excel_etiqueta": "Enganche fraccionado mas reserva",
            "excel_columna": "C",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": f"{sb['enganche_fraccionado']} + {sb['reserva']}",
            "formula_supabase": "engancheafraccionar + montoreserva",
            "notas": "Excel fila 33 incluye reserva; Supabase se suman ambos campos.",
        },
        {
            "concepto": "Monto a financiar",
            "excel_etiqueta": "Monto a financiar",
            "excel_columna": "H / I (USD)",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": sb["monto_financiar"],
            "formula_supabase": "Valor directo",
            "notas": "",
        },
        {
            "concepto": "Otros gastos",
            "excel_etiqueta": "Gastos y honorarios legales",
            "excel_columna": "C",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": f"{sb['gastos_compra']} + {sb['otros_cargos']}",
            "formula_supabase": "gastoscompranofinanc + otroscargosaddnofinanc",
            "notas": "En Excel solo gastoscompranofinanc (sin otros cargos).",
        },
        {
            "concepto": "Precio total con gastos",
            "excel_etiqueta": "Precio total",
            "excel_columna": "C",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": sb["precio_total_con_gastos"],
            "formula_supabase": "Valor directo",
            "notas": "Debe coincidir con la suma de los 5 conceptos anteriores.",
        },
        {
            "concepto": "Precio de venta (ref.)",
            "excel_etiqueta": "Precio de venta",
            "excel_columna": "H / I (USD)",
            "supabase_tabla": "vista_crm_tratos_cotizaciones",
            "supabase_columna": sb["precio_venta"],
            "formula_supabase": "Valor directo",
            "notas": "Solo referencia; no entra en la suma principal.",
        },
    ]
    return pd.DataFrame(rows)


def _supabase_tables_reference() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "tabla_vista": "vista_crm_tratos_cotizaciones",
                "uso": "Fuente principal de montos de la cotización",
                "columnas_clave": (
                    "id_crmtratocotizacion, nombreproyecto_producto, nombre_producto, "
                    "lote_producto, montoreserva, montopromesa, engancheafraccionar, "
                    "montofinanciar, gastoscompranofinanc, otroscargosaddnofinanc, "
                    "preciototalygastoscompranofin, preciototalventa"
                ),
            },
            {
                "tabla_vista": "crm_tratos_reservas",
                "uso": "Filtrar cotizaciones con reserva activa",
                "columnas_clave": (
                    "crmtratoreserva_crmtratocotizacion_id, isactive_crmtratoreserva, "
                    "estatus_crmtratoreserva"
                ),
            },
        ]
    )


def _join_rules_reference() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"regla": "Proyecto", "excel": "NEO2 (hojas del Excel)", "supabase": "nombreproyecto_producto = 'NEO 2'"},
            {"regla": "Unidad", "excel": "Número de apartamento en ficha", "supabase": "nombre_producto o lote_producto"},
            {"regla": "Cotización válida", "excel": "—", "supabase": "Debe tener reserva activa (isactive = true)"},
            {
                "regla": "Si hay varias cotizaciones",
                "excel": "—",
                "supabase": "Priorizar estatus Autorizada; si empatan, la más reciente",
            },
            {"regla": "Tolerancia", "excel": "±0.01", "supabase": "±0.01"},
            {
                "regla": "Fórmula total (Daniel)",
                "excel": "Reserva + PVC + Enganche + Financiar + Gastos = Precio total",
                "supabase": "Misma fórmula con columnas _crmtratocotizacion",
            },
        ]
    )


def _instrucciones_rows() -> list[list[str]]:
    return [
        ["REPORTE DE COMPARACIÓN — Excel vs Supabase (NEO2)", ""],
        ["", ""],
        ["¿Qué es este archivo?", ""],
        [
            "Compara cada apartamento del Excel '01-ESTADO DE CUENTA INTERNO' "
            "contra la cotización en Supabase que tiene reserva activa.",
            "",
        ],
        ["", ""],
        ["Hojas del reporte", "Para qué sirve"],
        ["1. Instrucciones", "Esta guía"],
        ["2. Resumen", "Totales y estado general"],
        ["3. Mapeo columnas", "Qué campo de Excel se compara con qué columna de Supabase"],
        ["4. Reglas Supabase", "Tablas, cruce y criterio de cotización válida"],
        ["5. Solo diferencias", "Lista accionable: solo conceptos que no cuadran (revisar primero)"],
        ["6. Por unidad", "Vista resumida: una fila por apartamento"],
        ["7. Detalle por unidad", "Desglose completo concepto por concepto"],
        ["8. Excel sin match", "Apartamentos del Excel sin cotización en Supabase"],
        ["", ""],
        ["¿Qué hacer si hay DIFERENCIA?", ""],
        ["Paso 1", "Ir a la hoja 'Solo diferencias' y ubicar el apartamento y concepto en rojo."],
        ["Paso 2", "Abrir la hoja 'Mapeo columnas' para ver qué columna de Supabase revisar."],
        ["Paso 3", "En Supabase, buscar la cotización por id_crmtratocotizacion (columna en el reporte)."],
        ["Paso 4", "Corregir el monto en Supabase o validar si el Excel está desactualizado."],
        ["Paso 5", "Volver a ejecutar la comparación en la aplicación."],
        ["", ""],
        ["Colores en el detalle", ""],
        ["Rojo", "El concepto no coincide entre Excel y Supabase"],
        ["Verde", "El concepto coincide (dentro de tolerancia ±0.01)"],
        ["Amarillo", "Falta dato en Excel o Supabase"],
    ]


def _resumen_rows(result: ComparisonResult) -> pd.DataFrame:
    s = result.resumen
    diff_units = [u for u in result.units if u.estado_unidad == "DIFERENCIA"]
    campos_diff = [r for r in result.rows if r.estado == "DIFERENCIA" and r.campo in DESGLOSE_LABELS]

    rows = [
        {"indicador": "Fecha de generación", "valor": datetime.now().strftime("%Y-%m-%d %H:%M")},
        {"indicador": "Proyecto", "valor": "NEO2"},
        {"indicador": "Fuente Excel", "valor": "01-ESTADO DE CUENTA INTERNO..xlsx (fichas por cliente)"},
        {"indicador": "Fuente Supabase", "valor": "vista_crm_tratos_cotizaciones + crm_tratos_reservas"},
        {"indicador": "", "valor": ""},
        {"indicador": "Unidades comparadas", "valor": s["unidades"]},
        {"indicador": "Unidades OK", "valor": s["unidades_ok"]},
        {"indicador": "Unidades con diferencia", "valor": s["unidades_con_diferencia"]},
        {"indicador": "Conceptos comparados (total filas)", "valor": s["total"]},
        {"indicador": "Conceptos OK", "valor": s["ok"]},
        {"indicador": "Conceptos con diferencia", "valor": s["diferencia"]},
        {"indicador": "Conceptos sin dato", "valor": s["sin_dato"]},
        {"indicador": "Excel sin cotización en Supabase", "valor": len(result.excel_sin_match)},
        {"indicador": "", "valor": ""},
        {
            "indicador": "Apartamentos con diferencia (lista)",
            "valor": ", ".join(f"Apto {u.unidad}" for u in diff_units[:30])
            + ("…" if len(diff_units) > 30 else ""),
        },
        {
            "indicador": "Conceptos que más fallan",
            "valor": _top_failing_fields(campos_diff),
        },
    ]
    return pd.DataFrame(rows)


def _top_failing_fields(rows: list[ComparisonRow]) -> str:
    if not rows:
        return "Ninguno"
    counts: dict[str, int] = {}
    for r in rows:
        counts[r.campo] = counts.get(r.campo, 0) + 1
    ordered = sorted(counts.items(), key=lambda x: -x[1])
    return ", ".join(f"{DESGLOSE_LABELS.get(k, k)} ({v})" for k, v in ordered[:5])


def _solo_diferencias(result: ComparisonResult) -> pd.DataFrame:
    rows: list[dict] = []
    for r in result.rows:
        if r.estado != "DIFERENCIA":
            continue
        rows.append(
            {
                "apartamento": r.unidad,
                "proyecto": r.proyecto,
                "id_cotizacion": r.id_crmtratocotizacion,
                "seccion": r.seccion,
                "concepto": DESGLOSE_LABELS.get(r.campo, r.campo),
                "valor_excel": r.valor_excel,
                "valor_supabase": r.valor_supabase,
                "diferencia": r.diferencia,
                "que_revisar": ACCION_POR_CAMPO.get(r.campo, "Revisar monto en Supabase y en la ficha Excel."),
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["apartamento", "seccion", "concepto"])
    return df


def _por_unidad_claro(result: ComparisonResult) -> pd.DataFrame:
    rows = []
    for u in result.units:
        rows.append(
            {
                "apartamento": u.unidad,
                "proyecto": u.proyecto,
                "estado": u.estado_unidad,
                "conceptos_ok": f"{u.ok}/{u.total_campos}",
                "conceptos_con_diferencia": u.diferencias,
                "sin_dato": u.sin_dato,
                "total_excel": u.total_excel,
                "total_supabase": u.total_supabase,
                "diferencia_total": (
                    round(u.total_excel - u.total_supabase, 2)
                    if u.total_excel is not None and u.total_supabase is not None
                    else None
                ),
                "id_cotizacion": u.id_crmtratocotizacion,
            }
        )
    return pd.DataFrame(rows)


def _detail_with_sections(result: ComparisonResult) -> pd.DataFrame:
    rows: list[dict] = []
    for unit in result.units:
        unit_rows = [r for r in result.rows if r.join_key == unit.join_key]
        rows.append(
            {
                "tipo_fila": "UNIDAD",
                "apartamento": unit.unidad,
                "proyecto": unit.proyecto,
                "seccion": "",
                "concepto": f"{unit.estado_unidad} — {unit.ok}/{unit.total_campos} conceptos OK",
                "valor_excel": unit.total_excel,
                "valor_supabase": unit.total_supabase,
                "diferencia": (
                    round(unit.total_excel - unit.total_supabase, 2)
                    if unit.total_excel is not None and unit.total_supabase is not None
                    else None
                ),
                "estado": unit.estado_unidad,
                "id_cotizacion": unit.id_crmtratocotizacion,
            }
        )
        for seccion, titulo in [
            ("desglose", "DESGLOSE FINANCIERO"),
            ("validacion", "VALIDACIÓN SUMA"),
            ("referencia", "REFERENCIA"),
        ]:
            sec_rows = [r for r in unit_rows if r.seccion == seccion]
            if not sec_rows:
                continue
            rows.append(
                {
                    "tipo_fila": "SECCION",
                    "apartamento": unit.unidad,
                    "proyecto": "",
                    "seccion": titulo,
                    "concepto": "",
                    "valor_excel": None,
                    "valor_supabase": None,
                    "diferencia": None,
                    "estado": "",
                    "id_cotizacion": "",
                }
            )
            for r in sec_rows:
                rows.append(
                    {
                        "tipo_fila": "CONCEPTO",
                        "apartamento": unit.unidad,
                        "proyecto": r.proyecto,
                        "seccion": "",
                        "concepto": DESGLOSE_LABELS.get(r.campo, r.campo),
                        "valor_excel": r.valor_excel,
                        "valor_supabase": r.valor_supabase,
                        "diferencia": r.diferencia,
                        "estado": r.estado,
                        "id_cotizacion": r.id_crmtratocotizacion,
                    }
                )
    return pd.DataFrame(rows)


def _style_header_row(ws, row: int = 1) -> None:
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _set_col_widths(ws, widths: dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def _format_money_cols(ws, col_indices: list[int], start_row: int = 2) -> None:
    for row in range(start_row, ws.max_row + 1):
        for col in col_indices:
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")


def _apply_estado_colors(ws, estado_col: int, start_row: int = 2, money_cols: list[int] | None = None) -> None:
    money_cols = money_cols or []
    for row in range(start_row, ws.max_row + 1):
        estado = ws.cell(row=row, column=estado_col).value
        if estado == "DIFERENCIA":
            for col in range(1, ws.max_column + 1):
                c = ws.cell(row=row, column=col)
                c.fill = RED_FILL
                if col == estado_col:
                    c.font = RED_FONT
        elif estado == "OK":
            ws.cell(row=row, column=estado_col).font = GREEN_FONT
            for col in money_cols:
                ws.cell(row=row, column=col).font = GREEN_FONT
        elif estado == "SIN_DATO":
            for col in range(1, ws.max_column + 1):
                c = ws.cell(row=row, column=col)
                c.fill = YELLOW_FILL


def _write_instrucciones(ws) -> None:
    for r_idx, row in enumerate(_instrucciones_rows(), start=1):
        ws.cell(row=r_idx, column=1, value=row[0])
        ws.cell(row=r_idx, column=2, value=row[1])
    ws.cell(row=1, column=1).font = TITLE_FONT
    for r in (3, 11, 20, 27):
        if r <= ws.max_row:
            ws.cell(row=r, column=1).font = SUBTITLE_FONT
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 70
    for r in range(1, ws.max_row + 1):
        ws.cell(row=r, column=1).alignment = WRAP
        ws.cell(row=r, column=2).alignment = WRAP


def export_report(result: ComparisonResult, output_path: Path | str) -> Path:
    output_path = Path(output_path)

    resumen = _resumen_rows(result)
    mapeo = _mapping_reference()
    tablas = _supabase_tables_reference()
    reglas = _join_rules_reference()
    diferencias = _solo_diferencias(result)
    por_unidad = _por_unidad_claro(result)
    detail = _detail_with_sections(result)
    sin_match = pd.DataFrame(
        {
            "clave_excel": result.excel_sin_match,
            "que_revisar": (
                "No hay cotización NEO2 con reserva activa para esta unidad. "
                "Verificar nombre_producto/lote_producto en Supabase o crear/activar reserva."
            ),
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Hoja 0: instrucciones (escrita manualmente)
        pd.DataFrame().to_excel(writer, sheet_name="Instrucciones", index=False)
        _write_instrucciones(writer.sheets["Instrucciones"])

        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        mapeo.to_excel(writer, sheet_name="Mapeo columnas", index=False)
        reglas.to_excel(writer, sheet_name="Reglas Supabase", index=False)
        tablas.to_excel(writer, sheet_name="Tablas Supabase", index=False)
        diferencias.to_excel(writer, sheet_name="Solo diferencias", index=False)
        por_unidad.to_excel(writer, sheet_name="Por unidad", index=False)
        detail.to_excel(writer, sheet_name="Detalle por unidad", index=False)
        if not sin_match.empty:
            sin_match.to_excel(writer, sheet_name="Excel sin match", index=False)

        # --- Resumen ---
        ws = writer.sheets["Resumen"]
        _style_header_row(ws)
        _set_col_widths(ws, {1: 38, 2: 55})
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=2).alignment = WRAP
        ws.freeze_panes = "A2"

        # --- Mapeo columnas ---
        ws = writer.sheets["Mapeo columnas"]
        _style_header_row(ws)
        _set_col_widths(ws, {1: 22, 2: 32, 3: 14, 4: 32, 5: 38, 6: 28, 7: 36})
        for r in range(2, ws.max_row + 1):
            for c in range(1, 8):
                ws.cell(row=r, column=c).alignment = WRAP
                ws.cell(row=r, column=c).border = BORDER
        ws.freeze_panes = "A2"

        # --- Reglas ---
        ws = writer.sheets["Reglas Supabase"]
        _style_header_row(ws)
        _set_col_widths(ws, {1: 22, 2: 28, 3: 50})
        for r in range(2, ws.max_row + 1):
            for c in range(1, 4):
                ws.cell(row=r, column=c).alignment = WRAP
                ws.cell(row=r, column=c).border = BORDER
        ws.freeze_panes = "A2"

        # --- Tablas ---
        ws = writer.sheets["Tablas Supabase"]
        _style_header_row(ws)
        _set_col_widths(ws, {1: 32, 2: 36, 3: 70})
        for r in range(2, ws.max_row + 1):
            for c in range(1, 4):
                ws.cell(row=r, column=c).alignment = WRAP
                ws.cell(row=r, column=c).border = BORDER

        # --- Solo diferencias ---
        ws = writer.sheets["Solo diferencias"]
        _style_header_row(ws)
        headers = [c.value for c in ws[1]]
        money_idx = [headers.index(h) + 1 for h in ("valor_excel", "valor_supabase", "diferencia") if h in headers]
        _format_money_cols(ws, money_idx)
        _set_col_widths(ws, {1: 12, 2: 10, 3: 38, 4: 14, 5: 36, 6: 14, 7: 14, 8: 14, 9: 55})
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).alignment = WRAP
                ws.cell(row=r, column=c).border = BORDER
        ws.freeze_panes = "A2"

        # --- Por unidad ---
        ws = writer.sheets["Por unidad"]
        _style_header_row(ws)
        headers = [c.value for c in ws[1]]
        estado_col = headers.index("estado") + 1 if "estado" in headers else None
        money_idx = [
            headers.index(h) + 1
            for h in ("total_excel", "total_supabase", "diferencia_total")
            if h in headers
        ]
        _format_money_cols(ws, money_idx)
        if estado_col:
            _apply_estado_colors(ws, estado_col, money_cols=money_idx)
        _set_col_widths(ws, {1: 12, 2: 10, 3: 14, 4: 14, 5: 18, 6: 10, 7: 14, 8: 14, 9: 14, 10: 38})
        ws.freeze_panes = "A2"

        # --- Detalle por unidad ---
        ws = writer.sheets["Detalle por unidad"]
        _style_header_row(ws)
        headers = [c.value for c in ws[1]]
        tipo_col = headers.index("tipo_fila") + 1 if "tipo_fila" in headers else None
        estado_col = headers.index("estado") + 1 if "estado" in headers else None
        money_idx = [
            headers.index(h) + 1
            for h in ("valor_excel", "valor_supabase", "diferencia")
            if h in headers
        ]
        _format_money_cols(ws, money_idx)

        for row_idx in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col).border = BORDER

            if tipo_col:
                tipo = ws.cell(row=row_idx, column=tipo_col).value
                if tipo == "UNIDAD":
                    for col in range(1, ws.max_column + 1):
                        c = ws.cell(row=row_idx, column=col)
                        c.fill = UNIT_FILL
                        c.font = UNIT_FONT
                elif tipo == "SECCION":
                    for col in range(1, ws.max_column + 1):
                        c = ws.cell(row=row_idx, column=col)
                        c.fill = SECTION_FILL
                        c.font = SECTION_FONT
                elif tipo == "CONCEPTO" and estado_col:
                    estado = ws.cell(row=row_idx, column=estado_col).value
                    if estado == "DIFERENCIA":
                        for col in range(1, ws.max_column + 1):
                            c = ws.cell(row=row_idx, column=col)
                            c.fill = RED_FILL
                            if col == estado_col:
                                c.font = RED_FONT
                    elif estado == "OK":
                        ws.cell(row=row_idx, column=estado_col).font = GREEN_FONT
                    elif estado == "SIN_DATO":
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row_idx, column=col).fill = YELLOW_FILL

        _set_col_widths(
            ws,
            {1: 10, 2: 12, 3: 10, 4: 22, 5: 42, 6: 14, 7: 14, 8: 14, 9: 12, 10: 38},
        )
        ws.freeze_panes = "A2"

        # --- Sin match ---
        if not sin_match.empty:
            ws = writer.sheets["Excel sin match"]
            _style_header_row(ws)
            _set_col_widths(ws, {1: 40, 2: 65})
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=2).alignment = WRAP
                ws.cell(row=r, column=2).fill = YELLOW_FILL

    return output_path
