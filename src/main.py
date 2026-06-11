"""Aplicación de escritorio: comparador Excel vs Supabase."""

from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))

from comparator import ComparisonResult, ComparisonRow, UnitSummary, compare_all
from config import ROOT_DIR, default_excel_path
from excel_loader import load_ficha_sheets
from report import export_report
from supabase_client import load_supabase_bundle

os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")

# Orden fijo del desglose financiero (Daniel)
DESGLOSE_CONCEPTOS: list[tuple[str, str]] = [
    ("reserva", "1. Reserva"),
    ("abono_pvc", "2. Abono PVC (promesa de compraventa)"),
    ("enganche_fraccionado", "3. Enganche fraccionado (+ reserva en fila 33 Excel)"),
    ("monto_financiar", "4. Monto a financiar"),
    ("otros_gastos", "5. Otros gastos (gastos legales en Excel)"),
    ("precio_total_con_gastos", "6. Precio total CON gastos (debe cuadrar con la suma)"),
]

VALIDACION_LABELS = {
    "suma_desglose": "Suma en Excel: Reserva+PVC+Enganche+Financiar+Gastos → vs Precio total Excel",
    "suma_desglose_supabase": "Suma en Supabase: mismos conceptos → vs preciototalygastoscompranofin",
}

REFERENCIA_LABELS = {
    "precio_venta": "Precio de venta (referencia, sin gastos)",
}


class ComparadorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Comparador Cotizaciones — Excel vs Supabase (NEO2)")
        self.geometry("1280x780")
        self.minsize(1000, 560)
        self.configure(bg="#f4f4f4")

        self.excel_path = tk.StringVar(value=str(default_excel_path()))
        self.filtro = tk.StringVar(value="diff")
        self.result: ComparisonResult | None = None
        self._build_ui()
        self._check_tk_version()

    def _check_tk_version(self) -> None:
        if tk.TkVersion < 8.6:
            messagebox.showerror("Tk incompatible", "Ejecuta: ./run.sh")
            self.destroy()

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 5}

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Archivo Excel:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.excel_path, width=55).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(top, text="Seleccionar…", command=self._pick_excel).grid(row=0, column=2, **pad)
        self.btn_compare = ttk.Button(top, text="Comparar", command=self._run_compare)
        self.btn_compare.grid(row=0, column=3, **pad)
        self.btn_export = ttk.Button(
            top, text="Exportar reporte", command=self._export, state="disabled"
        )
        self.btn_export.grid(row=0, column=4, **pad)

        ttk.Label(top, text="Mostrar:").grid(row=1, column=0, sticky="w", **pad)
        filtro_frame = ttk.Frame(top)
        filtro_frame.grid(row=1, column=1, columnspan=2, sticky="w", **pad)
        for val, label in [
            ("diff", "Solo unidades con diferencia"),
            ("all", "Todas las unidades"),
            ("ok", "Solo unidades OK"),
        ]:
            ttk.Radiobutton(
                filtro_frame,
                text=label,
                variable=self.filtro,
                value=val,
                command=self._refresh_tree,
            ).pack(side="left", padx=(0, 14))

        top.grid_columnconfigure(1, weight=1)

        # Panel guía
        guia = ttk.LabelFrame(self, text="¿Cómo leer el desglose?", padding=8)
        guia.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Label(
            guia,
            justify="left",
            text=(
                "Por cada APARTAMENTO se comparan 6 conceptos: Excel (ficha del cliente) vs Supabase (cotización con reserva activa).\n"
                "Fórmula Daniel: Reserva + Abono PVC + Enganche fraccionado + Monto a financiar + Otros gastos = Precio total con gastos.\n"
                "Rojo = el concepto no coincide entre Excel y Supabase. Expande cada unidad para ver el detalle."
            ),
        ).pack(anchor="w")

        self.status = ttk.Label(self, text="Listo. Pulsa Comparar.", anchor="w")
        self.status.pack(fill="x", padx=12, pady=(0, 4))

        # Resumen chips
        self.resumen_frame = ttk.Frame(self, padding=(12, 0))
        self.resumen_frame.pack(fill="x")
        self.lbl_unidades = ttk.Label(self.resumen_frame, text="", font=("TkDefaultFont", 11, "bold"))
        self.lbl_unidades.pack(side="left", padx=(0, 20))
        self.lbl_ok = ttk.Label(self.resumen_frame, text="", foreground="#1a7f37")
        self.lbl_ok.pack(side="left", padx=(0, 20))
        self.lbl_diff = ttk.Label(self.resumen_frame, text="", foreground="#c62828")
        self.lbl_diff.pack(side="left")

        table_frame = ttk.Frame(self, padding=(10, 6, 10, 10))
        table_frame.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Treeview", rowheight=28, font=("TkDefaultFont", 11))
        style.configure("Treeview.Heading", font=("TkDefaultFont", 11, "bold"))

        columns = ("concepto", "excel", "supabase", "diff", "estado")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", height=20)
        self.tree.heading("#0", text="Agrupación", anchor="w")
        self.tree.column("#0", width=280, stretch=True, minwidth=200)
        self.tree.heading("concepto", text="Concepto / Detalle")
        self.tree.heading("excel", text="Excel (ficha cliente)")
        self.tree.heading("supabase", text="Supabase (cotización)")
        self.tree.heading("diff", text="Diferencia")
        self.tree.heading("estado", text="Estado")
        for col, w in zip(columns, (340, 130, 130, 100, 90)):
            anchor = "e" if col in ("excel", "supabase", "diff") else ("center" if col == "estado" else "w")
            self.tree.column(col, width=w, anchor=anchor, stretch=col == "concepto")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("ok", foreground="#1a7f37")
        self.tree.tag_configure("diff", foreground="#c62828", font=("TkDefaultFont", 11, "bold"))
        self.tree.tag_configure("warn", foreground="#b8860b")
        self.tree.tag_configure("unit_ok", foreground="#1565c0", font=("TkDefaultFont", 12, "bold"))
        self.tree.tag_configure("unit_bad", foreground="#c62828", font=("TkDefaultFont", 12, "bold"))
        self.tree.tag_configure("section", foreground="#424242", font=("TkDefaultFont", 11, "bold"))
        self.tree.tag_configure("meta", foreground="#616161", font=("TkDefaultFont", 10, "italic"))

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            initialdir=str(ROOT_DIR),
        )
        if path:
            self.excel_path.set(path)

    def _set_status(self, text: str) -> None:
        self.status.configure(text=text)

    def _run_compare(self) -> None:
        self.btn_compare.configure(state="disabled")
        self.btn_export.configure(state="disabled")
        self._set_status("Comparando… (puede tardar ~1 min)")
        threading.Thread(target=self._compare_worker, daemon=True).start()

    def _compare_worker(self) -> None:
        try:
            excel_path = Path(self.excel_path.get())
            if not excel_path.exists():
                raise FileNotFoundError(f"No existe el archivo: {excel_path}")

            excel_df = load_ficha_sheets(excel_path)
            cot_df = load_supabase_bundle()
            result = compare_all(excel_df, cot_df)
            self.after(0, lambda: self._on_compare_done(result, None))
        except Exception as exc:
            self.after(0, lambda: self._on_compare_done(None, exc))

    def _on_compare_done(self, result: ComparisonResult | None, error: Exception | None) -> None:
        self.btn_compare.configure(state="normal")
        if error:
            messagebox.showerror("Error", str(error))
            self._set_status(f"Error: {error}")
            return

        assert result is not None
        self.result = result
        self._update_resumen_chips(result)
        self._refresh_tree()
        s = result.resumen
        self._set_status(
            f"Comparación lista — {s['unidades']} unidades · "
            f"{s['unidades_ok']} OK · {s['unidades_con_diferencia']} con diferencia · "
            f"{len(result.excel_sin_match)} sin match en Supabase"
        )
        self.btn_export.configure(state="normal")

    def _update_resumen_chips(self, result: ComparisonResult) -> None:
        s = result.resumen
        self.lbl_unidades.configure(text=f"Unidades comparadas: {s['unidades']}")
        self.lbl_ok.configure(text=f"✓ OK: {s['unidades_ok']}")
        self.lbl_diff.configure(text=f"✗ Con diferencia: {s['unidades_con_diferencia']}")

    def _should_show_unit(self, unit: UnitSummary) -> bool:
        f = self.filtro.get()
        if f == "all":
            return True
        if f == "ok":
            return unit.estado_unidad == "OK"
        return unit.estado_unidad == "DIFERENCIA"

    def _refresh_tree(self) -> None:
        if not self.result:
            return
        self._populate_tree(self.result)

    def _rows_by_campo(self, rows: list[ComparisonRow]) -> dict[str, ComparisonRow]:
        return {r.campo: r for r in rows}

    def _insert_concept_row(self, parent: str, label: str, row: ComparisonRow) -> None:
        tag = "ok"
        if row.estado == "DIFERENCIA":
            tag = "diff"
        elif row.estado == "SIN_DATO":
            tag = "warn"
        self.tree.insert(
            parent,
            "end",
            text="",
            values=(
                label,
                self._fmt(row.valor_excel),
                self._fmt(row.valor_supabase),
                self._fmt(row.diferencia),
                row.estado,
            ),
            tags=(tag,),
        )

    def _insert_section(self, parent: str, title: str) -> str:
        return self.tree.insert(
            parent,
            "end",
            text="",
            values=(title, "", "", "", ""),
            tags=("section",),
            open=True,
        )

    def _unit_header_text(self, unit: UnitSummary) -> str:
        icon = "✓" if unit.estado_unidad == "OK" else "✗"
        return f"{icon}  Apto {unit.unidad}  ·  {unit.proyecto}"

    def _unit_subtitle(self, unit: UnitSummary) -> str:
        cid = (unit.id_crmtratocotizacion or "")[:8]
        resumen = f"{unit.ok}/{unit.total_campos} conceptos OK"
        if unit.diferencias:
            resumen += f"  ·  {unit.diferencias} con diferencia"
        if unit.sin_dato:
            resumen += f"  ·  {unit.sin_dato} sin dato"
        totales = ""
        if unit.total_excel is not None and unit.total_supabase is not None:
            totales = (
                f"Total Excel {self._fmt(unit.total_excel)} vs "
                f"Supabase {self._fmt(unit.total_supabase)}"
            )
        return f"Cotización {cid}…  |  {resumen}  |  {totales}"

    def _populate_tree(self, result: ComparisonResult) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows_by_unit: dict[str, list[ComparisonRow]] = {}
        for row in result.rows:
            rows_by_unit.setdefault(row.join_key, []).append(row)

        for unit in result.units:
            if not self._should_show_unit(unit):
                continue

            u_tag = "unit_ok" if unit.estado_unidad == "OK" else "unit_bad"
            unit_node = self.tree.insert(
                "",
                "end",
                text=self._unit_header_text(unit),
                values=(
                    self._unit_subtitle(unit),
                    "",
                    "",
                    "",
                    unit.estado_unidad,
                ),
                tags=(u_tag,),
                open=unit.estado_unidad == "DIFERENCIA",
            )

            by_campo = self._rows_by_campo(rows_by_unit.get(unit.join_key, []))

            # Sección 1: Desglose
            sec_desglose = self._insert_section(
                unit_node,
                "▼  DESGLOSE FINANCIERO  —  comparar concepto por concepto",
            )
            for campo, label in DESGLOSE_CONCEPTOS:
                row = by_campo.get(campo)
                if row:
                    self._insert_concept_row(sec_desglose, label, row)

            # Referencia opcional
            for campo, label in REFERENCIA_LABELS.items():
                row = by_campo.get(campo)
                if row:
                    self._insert_concept_row(sec_desglose, label, row)

            # Sección 2: Validación suma
            sec_val = self._insert_section(
                unit_node,
                "▼  VALIDACIÓN SUMA  —  ¿los conceptos cuadran con el total?",
            )
            self.tree.insert(
                sec_val,
                "end",
                text="",
                values=(
                    "Fórmula: Reserva + PVC + Enganche + Financiar + Otros gastos = Precio total con gastos",
                    "",
                    "",
                    "",
                    "",
                ),
                tags=("meta",),
            )
            for campo, label in VALIDACION_LABELS.items():
                row = by_campo.get(campo)
                if row:
                    self._insert_concept_row(sec_val, label, row)

        # Sin match al final
        if result.excel_sin_match and self.filtro.get() in ("all", "diff"):
            sin_node = self.tree.insert(
                "",
                "end",
                text=f"⚠  Sin cotización en Supabase ({len(result.excel_sin_match)})",
                values=("Unidades del Excel sin reserva activa / sin match", "", "", "", ""),
                tags=("unit_bad",),
                open=True,
            )
            sec = self._insert_section(sin_node, "Unidades no encontradas en Supabase")
            for clave in result.excel_sin_match:
                self.tree.insert(
                    sec,
                    "end",
                    text="",
                    values=(clave, "", "", "", "SIN MATCH"),
                    tags=("warn",),
                )

    @staticmethod
    def _fmt(val: float | None) -> str:
        if val is None:
            return "—"
        return f"Q {val:,.2f}"

    def _export(self) -> None:
        if not self.result:
            return
        path = filedialog.asksaveasfilename(
            title="Guardar reporte",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="reporte_diferencias.xlsx",
            initialdir=str(ROOT_DIR),
        )
        if not path:
            return
        try:
            export_report(self.result, path)
            messagebox.showinfo("Exportado", f"Reporte guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al exportar", str(exc))


def main() -> None:
    app = ComparadorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
