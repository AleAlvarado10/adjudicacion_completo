"""
Hojas DECISIÓN POR ÍTEM y ESTRATEGIA FINALISTAS.

- Decisión por Ítem: una fila por ítem con quién gana en cada escenario, el monto
  de cada uno y el estado de homologación del ganador del E3.
- Estrategia Finalistas: cobertura y desempeño de cada proveedor, con una columna
  SÍ/NO para marcar finalistas; abajo, el resultado de consolidar en ese set.

Todo con fórmulas sobre 'Prop Adj Final', así que al marcar finalistas o llenar el
Escenario 3 los números se recalculan solos. La hoja Decisión por Ítem lleva cuatro
columnas auxiliares ocultas (L a O) que alimentan el bloque 2 de Estrategia.
"""

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .estilo import FONT_NAME, GRIS, NAVY
from .hojas_extra import _estilo_datos, _headers, _nota, _thin_border, _titulo

COL_ID = "ID SAP"
COL_PROV = "Nombre del proveedor"
COL_DEST = "País destino"
COL_DESC = "Descripción"
COL_CANT = "Cantidad Esperada"
COL_PUNT = "Puntaje Técnico\n(/5.0)"
COL_HOMOL = "Calidad /\nHomologación"
COL_BASELINE = "Baseline"

HOJA_DECISION = "Decisión por Ítem"
HOJA_ESTRATEGIA = "Estrategia Finalistas"

# Columnas auxiliares (ocultas) de la hoja Decisión por Ítem
AUX_CUBIERTO, AUX_COSTO_MIN, AUX_MONTO, AUX_BASE = 12, 13, 14, 15  # L, M, N, O


def _refs(columnas_prop, prop_rango):
    """Rangos absolutos de Prop Adj Final, por nombre de columna."""
    fila_ini, fila_fin = prop_rango
    hoja = "'Prop Adj Final'"

    def rng(nombre):
        for i, c in enumerate(columnas_prop, start=1):
            if str(c).strip() == str(nombre).strip():
                letra = get_column_letter(i)
                return f"{hoja}!${letra}${fila_ini}:${letra}${fila_fin}"
        return None

    return rng


def escribir_decision_por_item(wb, hoja_df, prop_rango, columnas_prop,
                                ronda_vigente=1, providers=(),
                                fila_estrategia_ini=6):
    """Una fila por ítem: quién gana en cada escenario y con qué monto."""
    rng = _refs(columnas_prop, prop_rango)
    r_id, r_prov, r_dest = rng(COL_ID), rng(COL_PROV), rng(COL_DEST)
    r_cant, r_homol, r_base = rng(COL_CANT), rng(COL_HOMOL), rng(COL_BASELINE)
    r_costo = rng(f"Costo total  (USD) R{ronda_vigente}") or rng(
        f"Costo total (USD) R{ronda_vigente}")
    pcts = {e: rng(f"% Adj. {e}") for e in ("E1", "E2")}
    pcts["E3"] = rng("% Adj. E3\n← MANUAL") or rng("% Adj. E3")
    montos = {e: rng(f"Monto {e}\n(USD)") for e in ("E1", "E2", "E3")}
    if not r_id or not r_prov:
        return None, None

    items = (hoja_df[[COL_ID, COL_DEST, COL_DESC]].drop_duplicates(subset=[COL_ID])
             if COL_DEST in hoja_df.columns else
             hoja_df[[COL_ID]].drop_duplicates().assign(**{COL_DEST: "", COL_DESC: ""}))

    ws = wb.create_sheet(HOJA_DECISION)
    n_cols = 10
    _titulo(ws, "DECISIÓN POR ÍTEM — qué proveedor queda en cada escenario", n_cols)
    _nota(ws, "Todo calculado con fórmulas sobre «Prop Adj Final». Las columnas L a O "
              "están ocultas: alimentan el bloque 2 de «Estrategia Finalistas».", n_cols)
    _headers(ws, ["ID SAP", "País", "Descripción", "E1 — Proveedor", "E2 — Proveedor",
                  "E3 — Proveedor", "Monto E1", "Monto E2", "Monto E3",
                  "Baseline"], row=3)

    f0 = 4
    providers = list(providers)
    fila_e_fin = fila_estrategia_ini + max(len(providers), 1) - 1
    flags = f"'{HOJA_ESTRATEGIA}'!$G${fila_estrategia_ini}:$G${fila_e_fin}"
    nombres = f"'{HOJA_ESTRATEGIA}'!$A${fila_estrategia_ini}:$A${fila_e_fin}"
    # 1 cuando la fila de Prop Adj Final es de un proveedor marcado como finalista
    es_finalista = f"COUNTIFS({nombres},{r_prov},{flags},\"SÍ\")"

    for i, (_, item) in enumerate(items.iterrows()):
        r = f0 + i
        ws.cell(row=r, column=1, value=item[COL_ID])
        ws.cell(row=r, column=2, value=item.get(COL_DEST))
        ws.cell(row=r, column=3, value=item.get(COL_DESC))
        for j, e in enumerate(("E1", "E2", "E3")):
            if pcts[e]:
                ws.cell(row=r, column=4 + j, value=(
                    f'=IFERROR(INDEX({r_prov},SUMPRODUCT(MATCH(1,'
                    f'({r_id}=$A{r})*({pcts[e]}>0),0))),"")'))
            if montos[e]:
                celda = ws.cell(row=r, column=7 + j,
                                value=f'=SUMIFS({montos[e]},{r_id},$A{r})')
                celda.number_format = "#,##0.00"
        if r_base:
            # Mismo Baseline que la hoja Prop Adj Final, traído por ID.
            celda = ws.cell(row=r, column=10, value=(
                f'=IFERROR(INDEX({r_base},MATCH($A{r},{r_id},0)),"")'))
            celda.number_format = "#,##0.00"

        # --- auxiliares para el set de finalistas ---
        if r_costo:
            ws.cell(row=r, column=AUX_CUBIERTO, value=(
                f'=IF(SUMPRODUCT(({r_id}=$A{r})*({r_costo}>0)*{es_finalista})>0,1,0)'))
            # Un término por proveedor: si está marcado, su costo mínimo para
            # este ítem; si no, un número enorme que MIN descarta. Sin arrays,
            # así funciona igual en Excel y en LibreOffice.
            terminos = []
            for k in range(len(providers)):
                fila_e = fila_estrategia_ini + k
                celda_nom = f"'{HOJA_ESTRATEGIA}'!$A${fila_e}"
                celda_flag = f"'{HOJA_ESTRATEGIA}'!$G${fila_e}"
                minifs = (f'_xlfn.MINIFS({r_costo},{r_id},$A{r},'
                          f'{r_prov},{celda_nom})')
                terminos.append(
                    f'IF({celda_flag}="SÍ",IFERROR(1/(1/{minifs}),9E+99),9E+99)')
            if terminos:
                expr = "MIN(" + ",".join(terminos) + ")"
                ws.cell(row=r, column=AUX_COSTO_MIN,
                        value=f'=IF({expr}>=9E+99,"",{expr})')
            if r_cant:
                ws.cell(row=r, column=AUX_MONTO, value=(
                    f'=IF(M{r}="",0,M{r}*IFERROR(AVERAGEIFS({r_cant},{r_id},$A{r}),0))'))
            if r_base and r_cant:
                ws.cell(row=r, column=AUX_BASE, value=(
                    f'=IFERROR(AVERAGEIFS({r_base},{r_id},$A{r})*'
                    f'AVERAGEIFS({r_cant},{r_id},$A{r}),0)'))

    _estilo_datos(ws, len(items), n_cols, header_row=3)
    for c in range(AUX_CUBIERTO, AUX_BASE + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 46
    for c in range(4, 11):
        ws.column_dimensions[get_column_letter(c)].width = 20
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    return ws, (f0, f0 + len(items) - 1)


def escribir_estrategia_finalistas(wb, providers, prop_rango, columnas_prop,
                                    decision_rango, ronda_vigente=1,
                                    total_items=0, fila_ini=6):
    """Cobertura por proveedor y resultado de consolidar en los finalistas."""
    rng = _refs(columnas_prop, prop_rango)
    r_prov, r_punt = rng(COL_PROV), rng(COL_PUNT)
    r_costo = rng(f"Costo total  (USD) R{ronda_vigente}") or rng(
        f"Costo total (USD) R{ronda_vigente}")
    pct_e3 = rng("% Adj. E3\n← MANUAL") or rng("% Adj. E3")
    monto_e3 = rng("Monto E3\n(USD)")
    if not r_prov:
        return None

    dec_ini, dec_fin = decision_rango
    dec = f"'{HOJA_DECISION}'"
    n_cols = 7
    ws = wb.create_sheet(HOJA_ESTRATEGIA)

    _titulo(ws, "ESTRATEGIA DE PROVEEDORES — ¿con cuántos nos quedamos?", n_cols)
    _nota(ws, "Marca SÍ / NO en la columna «¿Finalista?» y el bloque 2 recalcula "
              "cobertura, monto y ahorro de ese set.", n_cols)

    ws.cell(row=4, column=1, value="1. COBERTURA POR PROVEEDOR").font = Font(
        name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
    _headers(ws, ["Proveedor", "Cobertura de cotización", "Ítems con téc. ≥ 2.0",
                  "Téc. prom.", "Ítems adj. en E3", "Monto que tiene hoy en E3",
                  "¿Finalista?"], row=5)

    provs = sorted(providers)
    for i, prov in enumerate(provs):
        r = fila_ini + i
        ws.cell(row=r, column=1, value=prov)
        if r_costo:
            ws.cell(row=r, column=2, value=(
                f'=IF({total_items}=0,"",COUNTIFS({r_prov},$A{r},{r_costo},">0")'
                f'/{total_items})'))
            ws.cell(row=r, column=2).number_format = "0.0%"
        if r_punt:
            ws.cell(row=r, column=3, value=f'=COUNTIFS({r_prov},$A{r},{r_punt},">=2")')
            ws.cell(row=r, column=4, value=(
                f'=IFERROR(AVERAGEIFS({r_punt},{r_prov},$A{r},{r_punt},">0"),"")'))
            ws.cell(row=r, column=4).number_format = "0.00"
        if pct_e3:
            ws.cell(row=r, column=5, value=f'=COUNTIFS({r_prov},$A{r},{pct_e3},">0")')
        if monto_e3:
            ws.cell(row=r, column=6, value=f'=SUMIFS({monto_e3},{r_prov},$A{r})')
            ws.cell(row=r, column=6).number_format = "#,##0.00"
        ws.cell(row=r, column=7, value="NO")

    _estilo_datos(ws, len(provs), n_cols, header_row=5, input_cols=(7,))
    fila_fin = fila_ini + len(provs) - 1

    # ── Bloque 2: resultado del set marcado ──────────────────────────────────
    f = fila_fin + 2
    ws.cell(row=f, column=1, value="2. RESULTADO DEL SET SELECCIONADO").font = Font(
        name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
    f += 1
    cub = f"{dec}!$L${dec_ini}:$L${dec_fin}"
    monto = f"{dec}!$N${dec_ini}:$N${dec_fin}"
    base = f"{dec}!$O${dec_ini}:$O${dec_fin}"
    monto_e3_total = f'=SUM({monto_e3})' if monto_e3 else "0"

    # TEXTJOIN con un argumento escalar por proveedor: así no necesita fórmula
    # matricial y funciona igual en Excel y en LibreOffice.
    args = ",".join(f'IF($G${r}="SÍ",$A${r},"")'
                    for r in range(fila_ini, fila_fin + 1))
    hay = f'COUNTIF($G${fila_ini}:$G${fila_fin},"SÍ")'

    filas = [
        ("N° de proveedores finalistas", f'={hay}',
         f'=IF({hay}=0,"— marca SÍ en la columna de arriba —",'
         f'_xlfn.TEXTJOIN(" + ",TRUE,{args}))', "0"),
        ("Ítems cubiertos", f'=IF({hay}=0,"",SUM({cub}))',
         f'=IF(OR({hay}=0,{total_items}=0),"",SUM({cub})/{total_items})', "0"),
        ("Ítems SIN cubrir", f'=IF({hay}=0,"",{total_items}-SUM({cub}))',
         '=IF(B{f}="","",IF(B{f}=0,"Cobertura completa",'
         '"Faltan proveedores para esos ítems"))', "0"),
        ("Monto adjudicado", f'=IF({hay}=0,"",SUM({monto}))', None, "#,##0.00"),
        ("Ahorro vs baseline", f'=IF({hay}=0,"",SUM({base})-SUM({monto}))',
         f'=IF(OR({hay}=0,SUM({base})=0),"",(SUM({base})-SUM({monto}))/SUM({base}))',
         "#,##0.00"),
        ("Diferencia vs E3 actual",
         f'=IF({hay}=0,"",SUM({monto})-{monto_e3_total[1:]})', None, "#,##0.00"),
    ]
    for i, (etiqueta, formula, extra, fmt) in enumerate(filas):
        r = f + i
        ws.cell(row=r, column=1, value=etiqueta).font = Font(name=FONT_NAME, size=9, bold=True)
        ws.cell(row=r, column=2, value=formula)
        if fmt != "0":
            ws.cell(row=r, column=2).number_format = fmt
        if extra:
            ws.cell(row=r, column=3, value=extra.replace("{f}", str(f + 2)))
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = _thin_border()

    ws.cell(row=f + 1, column=3).number_format = "0.0%"
    ws.cell(row=f + 4, column=3).number_format = "0.0%"
    ws.cell(row=f + 5, column=3, value=(
        f'=IF(B{f + 5}="","",IF(B{f + 5}>0,"Consolidar CUESTA más que el E3 actual",'
        '"Consolidar AHORRA respecto del E3 actual"))'))

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 40
    for c in range(4, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 20
    ws.sheet_view.showGridLines = False
    return ws
