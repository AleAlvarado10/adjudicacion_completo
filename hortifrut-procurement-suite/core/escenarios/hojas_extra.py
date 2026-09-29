"""
Escribe y da estilo a las hojas auxiliares del workbook final, replicando
la estructura y fórmulas del archivo de referencia:

  Léeme | Consolidado | Detalle Ofertas | Detalle Evaluación Técnica |
  Evaluación Económica | Plan Homologación | Prop Adj Final

Cada función devuelve la info que las hojas siguientes necesitan para
poder cruzarse por fórmula (ej. cuántas filas tiene 'Detalle Evaluación
Técnica', para que 'Plan Homologación' y 'Prop Adj Final' apunten al
rango correcto).
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

from .evaluacion import UMBRAL_ALTO_DEFAULT, UMBRAL_MEDIO_DEFAULT


def _fmt_inflacion(v):
    """El índice puede venir como número único o como texto ('según país')."""
    if isinstance(v, (int, float)):
        return f"{v:.2%}"
    return str(v)

NAVY = "FF1F4E79"
WHITE = "FFFFFFFF"
PARAM_BG = "FFD6DCE4"
INPUT_BG = "FFFFF2CC"
VERDE = "FFC6EFCE"
AMARILLO = "FFFFEB9C"
ROJO = "FFFFC7CE"
GRIS = "FFE7E6E6"
FONT_NAME = "Calibri"


def _thin_border():
    thin = Side(style="thin", color="FFB7B7B7")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _titulo(ws, texto, n_cols, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(n_cols, 1))
    cell = ws.cell(row=row, column=1, value=texto)
    cell.font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for c in range(1, n_cols + 1):
        cc = ws.cell(row=row, column=c)
        cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cc.border = _thin_border()
    ws.row_dimensions[row].height = 24


def _nota(ws, texto, n_cols, row=2):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(n_cols, 1))
    cell = ws.cell(row=row, column=1, value=texto)
    cell.font = Font(name=FONT_NAME, size=8, italic=True, color="FF595959")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _headers(ws, headers, row):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(30, len(str(h)) + 4))
    ws.row_dimensions[row].height = 30


def _estilo_datos(ws, n_rows, n_cols, header_row, input_cols=()):
    for r in range(header_row + 1, header_row + 1 + n_rows):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT_NAME, size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = _thin_border()
            if c in input_cols:
                cell.fill = PatternFill(start_color=INPUT_BG, end_color=INPUT_BG, fill_type="solid")


def _cf_texto(ws, col_letter, row_ini, row_fin, mapa_color):
    rng = f"{col_letter}{row_ini}:{col_letter}{row_fin}"
    for valor, color in mapa_color.items():
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{valor}"'],
                             fill=PatternFill(start_color=color, end_color=color, fill_type="solid"))
        )


# ---------------------------------------------------------------------------
def escribir_leeme(wb, ronda_actual, umbral_alto, umbral_medio, factor_modo,
                    providers, indice_inflacion=0.0, rondas=(1,)):
    ws = wb.create_sheet("Léeme")
    ws.column_dimensions["A"].width = 110
    filas = [
        "Propuesta Adjudicación + Plan Homologación — generado con la app de Streamlit",
        "",
        "Umbrales de homologación",
        f"🟢 ALTO ≥ {umbral_alto} -> HOMOLOGADO   |   🟡 MEDIO {umbral_medio}–{umbral_alto} -> POR REVISAR   |   "
        f"🔴 BAJO < {umbral_medio} -> NO HOMOLOGADO   |   sin puntaje -> S/DATOS. "
        "Editable en 'Detalle Evaluación Técnica' B2/D2.",
        "",
        "Ajuste EXW -> DDP (Factor Logístico)" if factor_modo == "tabla"
        else "Sin Factor Logístico",
        (
            "Costo total (USD) R{n} = Precio Unitario R{n} × (1 + Factor Logistico). "
            "Si el proveedor YA incluye el costo logístico en su precio, INCOTERM = DDP y "
            "Factor = 0 (no se ajusta). Si no lo incluye, INCOTERM = EXW y se aplica el factor "
            "de la ruta Proveedor + País destino, según la tabla de Factor Logístico que subiste. "
            "Proveedores sin ruta en la tabla quedan con Factor = 0 (no se inventa un ajuste)."
            if factor_modo == "tabla" else
            "No se subió tabla de Factor Logístico, así que este archivo NO trae las columnas "
            "INCOTERM, Factor Logistico ni 'Precio Histórico 2526 + factor'. "
            "Costo total R{n} = Precio Unitario R{n}, y el Baseline se calcula directo sobre el "
            "Precio Histórico. Si quieres el ajuste logístico, vuelve a generar el archivo "
            "subiendo la tabla en la app."
        ),
        "",
        "Precio Histórico / Presupuesto",
        "Precio Histórico se cruza por ID contra la hoja 'Evaluación Económica' (complétala a mano, o "
        "sube la de una ronda anterior para recuperar lo ya ingresado). Precio Presupuesto queda en "
        "blanco para completar a mano; mientras esté vacío, los semáforos vs Presupuesto muestran 'S/BASE'.",
        "",
        "Historial de Rondas y Ahorro",
        f"Esta es la Ronda {ronda_actual}. Columnas presentes: "
        + ", ".join(f"R{n}" for n in rondas)
        + ". El precio unitario de cada ronda se guarda por ID+Proveedor y se recupera "
        "automáticamente en las rondas siguientes. La ronda vigente (la última) es la que "
        "alimenta Mínimo costo ofertado, Desviación, Escenarios, Monto y Ahorro.",
        "",
        "Baseline",
        (
            ("Baseline = Precio Histórico × (1 + Factor Logistico) × (1 + Índice de inflación). "
             if factor_modo == "tabla" else
             "Baseline = Precio Histórico × (1 + Índice de inflación). ")
            + f"Índice de inflación usado: {_fmt_inflacion(indice_inflacion)}. Si el ítem no tiene Precio "
            "Histórico, el Baseline cae al Costo total de la Ronda 1. "
            "Ahorro_E1/E2/E3 = (Baseline − Costo total ronda vigente) × Volumen del escenario "
            "(en rojo si sale negativo, o sea si el precio subió)."
        ),
        "",
        "Escenarios",
        "Escenario 1 (automático) = 100% al proveedor más barato del ítem. "
        "Escenario 2 — MEJOR PUNTAJE TÉCNICO (automático) = 100% al proveedor con el MAYOR "
        "puntaje técnico del ítem, sin importar si ese puntaje es alto, medio o bajo (no filtra "
        "por homologación). Si dos proveedores empatan en puntaje, gana el más barato de los dos; "
        "si el ítem no tiene ningún puntaje, queda en 0%. "
        "Escenario 3 (manual) = tú decides el % por fila, parte en 0%.",
        "",
        "Proveedores detectados",
        "Proveedor: " + ", ".join(providers),
    ]
    for i, texto in enumerate(filas, start=1):
        cell = ws.cell(row=i, column=1, value=texto if texto else None)
        if texto and (texto.endswith("homologación") or texto.endswith("(Factor Logístico)")
                      or texto == "Sin Factor Logístico"
                      or texto.endswith("Presupuesto") or texto.endswith("Ahorro")
                      or texto == "Escenarios" or texto == "Proveedores detectados"
                      or texto == "Baseline"):
            cell.font = Font(name=FONT_NAME, size=10, bold=True)
        elif i == 1:
            cell.font = Font(name=FONT_NAME, size=12, bold=True)
        else:
            cell.font = Font(name=FONT_NAME, size=9)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    return ws


def escribir_consolidado_echo(wb, df_consolidado):
    ws = wb.create_sheet("Consolidado")
    for c, col_name in enumerate(df_consolidado.columns, start=1):
        cell = ws.cell(row=1, column=c, value=col_name)
        cell.font = Font(name=FONT_NAME, size=8, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = 14
    for r, (_, row) in enumerate(df_consolidado.iterrows(), start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=(None if pd_isna(val) else val))
            cell.font = Font(name=FONT_NAME, size=8)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    return ws


def pd_isna(v):
    try:
        import pandas as pd
        return bool(pd.isna(v))
    except Exception:
        return v is None


def escribir_detalle_ofertas(wb, df_ofertas):
    ws = wb.create_sheet("Detalle Ofertas")
    n_cols = len(df_ofertas.columns)
    _titulo(ws, 'DETALLE OFERTAS — Consolidado "explotado" a 1 fila por ítem + proveedor', n_cols)
    _nota(ws, 'Fuente: hoja "Consolidado". Juega el rol de "Elementos y servicios" del archivo original.', n_cols)
    _headers(ws, list(df_ofertas.columns), row=3)
    for r, (_, row) in enumerate(df_ofertas.iterrows(), start=4):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=(None if pd_isna(val) else val))
    _estilo_datos(ws, len(df_ofertas), n_cols, header_row=3)
    ws.freeze_panes = "A4"
    return ws


def escribir_detalle_evaluacion_tecnica(wb, df_tecnica, umbral_alto=UMBRAL_ALTO_DEFAULT,
                                         umbral_medio=UMBRAL_MEDIO_DEFAULT):
    """
    Columnas fijas: A=Clave, B=ID, C=Proveedor, D=Puntaje Técnico (/5.0, INPUT),
    E=Semáforo (fórmula), F=Estado Homologación (fórmula).
    Devuelve (ws, primera_fila_datos, ultima_fila_datos).
    """
    headers = ["Clave (ID|Proveedor)", "ID", "Proveedor", "Puntaje Técnico (/5.0)",
               "Semáforo", "Estado Homologación", "Estado desde Plan\n(importado)"]
    n_cols = len(headers)
    ws = wb.create_sheet("Detalle Evaluación Técnica")
    _titulo(ws, "DETALLE EVALUACIÓN TÉCNICA — por ítem + proveedor", n_cols)

    ws.cell(row=2, column=1, value="Umbral ALTO (homologa) ≥")
    ws.cell(row=2, column=2, value=umbral_alto)
    ws.cell(row=2, column=3, value="Umbral MEDIO (por revisar) ≥")
    ws.cell(row=2, column=4, value=umbral_medio)
    for c in range(1, 5):
        cell = ws.cell(row=2, column=c)
        cell.font = Font(name=FONT_NAME, size=8, bold=True)
        cell.fill = PatternFill(start_color=PARAM_BG, end_color=PARAM_BG, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
    ws.cell(row=2, column=6, value="Estado desde Plan manda sobre el puntaje →")
    c6 = ws.cell(row=2, column=6)
    c6.font = Font(name=FONT_NAME, size=7, italic=True, color="FF595959")
    c6.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    for c in range(5, n_cols + 1):
        ws.cell(row=2, column=c).border = _thin_border()

    _headers(ws, headers, row=3)

    data_start = 4
    last_row = data_start + len(df_tecnica) - 1
    for i, (_, row) in enumerate(df_tecnica.iterrows()):
        r = data_start + i
        ws.cell(row=r, column=1, value=row["Clave"])
        ws.cell(row=r, column=2, value=row.get("ID"))
        ws.cell(row=r, column=3, value=row.get("Proveedor"))
        puntaje = row.get("Puntaje Técnico (/5.0)")
        ws.cell(row=r, column=4, value=(None if pd_isna(puntaje) else puntaje))
        ws.cell(row=r, column=5, value=(
            f'=IF(D{r}="","S/DATOS",IF(D{r}>=$B$2,"🟢 ALTO",IF(D{r}>=$D$2,"🟡 MEDIO","🔴 BAJO")))'
        ))
        # El estado importado del Plan de Homologación tiene prioridad; si no
        # existe, se deriva del puntaje técnico y los umbrales.
        ws.cell(row=r, column=6, value=(
            f'=IF(G{r}<>"",G{r},'
            f'IF(D{r}="","S/DATOS",IF(D{r}>=$B$2,"HOMOLOGADO",IF(D{r}>=$D$2,"POR REVISAR","NO HOMOLOGADO"))))'
        ))
        estado_plan = row.get("Estado desde Plan")
        ws.cell(row=r, column=7, value=(None if pd_isna(estado_plan) else estado_plan))

    _estilo_datos(ws, len(df_tecnica), n_cols, header_row=3, input_cols=(4, 7))
    _cf_texto(ws, "F", data_start, last_row,
              {"HOMOLOGADO": VERDE, "POR REVISAR": AMARILLO, "NO HOMOLOGADO": ROJO, "S/DATOS": GRIS})
    _cf_texto(ws, "G", data_start, last_row,
              {"HOMOLOGADO": VERDE, "POR REVISAR": AMARILLO, "NO HOMOLOGADO": ROJO})
    _cf_texto(ws, "E", data_start, last_row,
              {"🟢 ALTO": VERDE, "🟡 MEDIO": AMARILLO, "🔴 BAJO": ROJO, "S/DATOS": GRIS})
    ws.freeze_panes = "A4"
    return ws, data_start, last_row


def escribir_evaluacion_economica(wb, df_economica):
    """Columnas fijas: A=ID, B=Descripción, C=Precio Histórico (crudo, INPUT), D=(limpio, fórmula)."""
    headers = ["ID", "Descripción", "Precio Histórico (crudo)", "Precio Histórico (limpio)"]
    n_cols = len(headers)
    ws = wb.create_sheet("Evaluación Económica")
    _titulo(ws, "EVALUACIÓN ECONÓMICA — precio de referencia por ID (ronda anterior)", n_cols)
    ws.row_dimensions[2].height = 8
    _headers(ws, headers, row=3)

    data_start = 4
    last_row = data_start + len(df_economica) - 1
    for i, (_, row) in enumerate(df_economica.iterrows()):
        r = data_start + i
        ws.cell(row=r, column=1, value=row["ID"])
        ws.cell(row=r, column=2, value=row.get("Descripcion"))
        crudo = row.get("Precio Histórico (crudo)", 0)
        ws.cell(row=r, column=3, value=(0 if pd_isna(crudo) else crudo))
        ws.cell(row=r, column=4, value=f'=IF(OR(C{r}="",C{r}=0),"",C{r})')

    _estilo_datos(ws, len(df_economica), n_cols, header_row=3, input_cols=(3,))
    ws.column_dimensions["B"].width = 32
    ws.freeze_panes = "A4"
    return ws, data_start, last_row


def escribir_plan_homologacion(wb, df_plan, providers, tecnica_rango):
    """tecnica_rango = (primera_fila, ultima_fila) de 'Detalle Evaluación Técnica'."""
    headers = ["ID", "Descripción"] + providers
    n_cols = len(headers)
    ws = wb.create_sheet("Plan Homologación")
    _titulo(ws, "PLAN HOMOLOGACIÓN — matriz ID x Proveedor (fuente: Detalle Evaluación Técnica)", n_cols)
    _nota(ws, "Se recalcula sola si cambian los puntajes de la hoja de Evaluación Técnica.", n_cols)
    _headers(ws, headers, row=3)

    t_ini, t_fin = tecnica_rango
    data_start = 4
    for i, (_, row) in enumerate(df_plan.iterrows()):
        r = data_start + i
        ws.cell(row=r, column=1, value=row["ID"])
        ws.cell(row=r, column=2, value=row.get("Descripcion"))
        for j, prov in enumerate(providers):
            c = 3 + j
            col_letter = get_column_letter(c)
            ws.cell(row=r, column=c, value=(
                f"=IFERROR(INDEX('Detalle Evaluación Técnica'!$F${t_ini}:$F${t_fin},"
                f"MATCH($A{r}&\"|\"&{col_letter}$3,'Detalle Evaluación Técnica'!$A${t_ini}:$A${t_fin},0)),"
                f'"S/DATOS")'
            ))

    _estilo_datos(ws, len(df_plan), n_cols, header_row=3)
    for j in range(len(providers)):
        col_letter = get_column_letter(3 + j)
        _cf_texto(ws, col_letter, data_start, data_start + len(df_plan) - 1,
                  {"HOMOLOGADO": VERDE, "POR REVISAR": AMARILLO, "NO HOMOLOGADO": ROJO, "S/DATOS": GRIS})
    ws.column_dimensions["B"].width = 32
    ws.freeze_panes = "C4"
    return ws


# ---------------------------------------------------------------------------
# Hoja RESUMEN — lectura ejecutiva de la propuesta
# ---------------------------------------------------------------------------

def escribir_resumen(wb, providers, prop_rango, columnas_prop, ronda_actual=1,
                     nombre_proceso=""):
    """
    Hoja RESUMEN: escenarios en columnas (E1/E2/E3), indicadores en filas.

    Tres bloques:
      1. Resumen por escenario  — ítems, % adjudicación, volumen, monto y ahorro
      2. Monto adjudicado por proveedor
      3. Ahorro por proveedor

    Todo con fórmulas sobre 'Prop Adj Final': si se edita el Escenario 3 o se
    corrige un puntaje técnico, el resumen se actualiza al abrir el archivo.
    """
    fila_ini, fila_fin = prop_rango
    hoja = "'Prop Adj Final'"

    def col(nombre):
        for i, c in enumerate(columnas_prop, start=1):
            if str(c).strip() == str(nombre).strip():
                return get_column_letter(i)
        return None

    def rng(nombre):
        letra = col(nombre)
        return f"{hoja}!${letra}${fila_ini}:${letra}${fila_fin}" if letra else None

    r_prov = rng("Nombre del proveedor")
    r_id = rng("ID SAP")

    escenarios = [
        ("E1 — más económico", rng("% Adj. E1"), rng("Vol. E1"),
         rng("Monto E1\n(USD)"), rng("Ahorro_E1")),
        ("E2 — mayor puntaje técnico", rng("% Adj. E2"), rng("Vol. E2"),
         rng("Monto E2\n(USD)"), rng("Ahorro_E2")),
        ("E3 — manual", rng("% Adj. E3\n← MANUAL"), rng("Vol. E3"),
         rng("Monto E3\n(USD)"), rng("Ahorro_E3")),
    ]
    escenarios = [e for e in escenarios if e[1] and e[3]]
    n_cols = 1 + len(escenarios)

    ws = wb.create_sheet("Resumen", 1)  # justo después de Léeme
    _titulo(ws, f"RESUMEN POR ESCENARIO — {nombre_proceso or 'proceso'} · Ronda "
                f"{ronda_actual} (USD)", n_cols)
    _nota(ws, "Todo calculado con fórmulas sobre «Prop Adj Final»: si cambias el "
              "Escenario 3 o un puntaje técnico, este resumen se actualiza solo.", n_cols)

    # ── Bloque 1: indicadores por escenario ──────────────────────────────────
    fila_hdr = 3
    _headers(ws, ["Indicador"] + [e[0] for e in escenarios], row=fila_hdr)

    indicadores = ["Ítems adjudicados", "% Adjudicación (de los ítems ofertados)",
                   "Volumen adjudicado", "Monto adjudicado (USD)",
                   "Ahorro total (USD)", "Ahorro sobre baseline [%]"]
    f_items = fila_hdr + 1
    for i, nombre in enumerate(indicadores):
        ws.cell(row=f_items + i, column=1, value=nombre)

    for j, (_, r_pct, r_vol, r_monto, r_ahorro) in enumerate(escenarios):
        c = 2 + j
        letra = get_column_letter(c)
        ws.cell(row=f_items, column=c, value=f'=COUNTIFS({r_pct},">0")')
        ws.cell(row=f_items + 1, column=c,
                value=f'=IF(COUNTA({r_id})=0,"",{letra}{f_items}/COUNTA({r_id}))')
        ws.cell(row=f_items + 2, column=c, value=f'=SUM({r_vol})' if r_vol else None)
        ws.cell(row=f_items + 3, column=c, value=f'=SUM({r_monto})')
        ws.cell(row=f_items + 4, column=c, value=f'=SUM({r_ahorro})' if r_ahorro else None)
        ws.cell(row=f_items + 5, column=c, value=(
            f'=IF(({letra}{f_items + 3}+{letra}{f_items + 4})=0,"",'
            f'{letra}{f_items + 4}/({letra}{f_items + 3}+{letra}{f_items + 4}))'
            if r_ahorro else None))

    _estilo_datos(ws, len(indicadores), n_cols, header_row=fila_hdr)
    for i, fmt in enumerate(["#,##0", "0.0%", "#,##0", "#,##0.00", "#,##0.00", "0.0%"]):
        for c in range(2, n_cols + 1):
            ws.cell(row=f_items + i, column=c).number_format = fmt

    # ── Bloques 2 y 3: por proveedor ─────────────────────────────────────────
    def bloque_proveedores(fila_titulo, titulo, indice_rango, formato, con_semaforo=False):
        ws.cell(row=fila_titulo, column=1, value=titulo).font = Font(
            name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
        fila_h = fila_titulo + 1
        _headers(ws, ["Proveedor"] + [e[0].split(" — ")[0] for e in escenarios], row=fila_h)
        f0 = fila_h + 1
        for i, prov in enumerate(sorted(providers)):
            r = f0 + i
            ws.cell(row=r, column=1, value=prov)
            for j, esc in enumerate(escenarios):
                rango = esc[indice_rango]
                ws.cell(row=r, column=2 + j,
                        value=f'=SUMIFS({rango},{r_prov},$A{r})' if rango else None)
        _estilo_datos(ws, len(providers), n_cols, header_row=fila_h)

        r_total = f0 + len(providers)
        ws.cell(row=r_total, column=1, value="TOTAL").font = Font(
            name=FONT_NAME, size=9, bold=True)
        for j in range(len(escenarios)):
            letra = get_column_letter(2 + j)
            celda = ws.cell(row=r_total, column=2 + j,
                            value=f"=SUM({letra}{f0}:{letra}{r_total - 1})")
            celda.font = Font(name=FONT_NAME, size=9, bold=True)
        for c in range(1, n_cols + 1):
            cc = ws.cell(row=r_total, column=c)
            cc.border = _thin_border()
            cc.fill = PatternFill(start_color=GRIS, end_color=GRIS, fill_type="solid")
        for r in range(f0, r_total + 1):
            for c in range(2, n_cols + 1):
                ws.cell(row=r, column=c).number_format = formato
        if con_semaforo:
            ws.conditional_formatting.add(
                f"B{f0}:{get_column_letter(n_cols)}{r_total}",
                CellIsRule(operator="lessThan", formula=["0"],
                           fill=PatternFill(start_color=ROJO, end_color=ROJO,
                                            fill_type="solid")),
            )
        return r_total

    f2 = f_items + len(indicadores) + 1
    fin2 = bloque_proveedores(f2, "MONTO ADJUDICADO POR PROVEEDOR (USD)", 3, "#,##0.00")
    fin3 = bloque_proveedores(fin2 + 2, "AHORRO POR PROVEEDOR (USD)", 4, "#,##0.00",
                              con_semaforo=True)

    ws.column_dimensions["A"].width = 38
    for c in range(2, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 24
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B4"
    return ws
