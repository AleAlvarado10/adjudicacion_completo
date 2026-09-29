"""
Hoja RESUMEN EJECUTIVO — la que se lleva a la reunión de adjudicación.

Tres bloques:
  1. Monto adjudicado por destino, con el proveedor principal del Escenario 3
  2. Qué proveedor queda en cada destino, con subtotales
  3. Ítems a decidir: aquellos donde el más económico (E1) y el de mejor puntaje
     técnico (E2) NO coinciden, con una columna para marcar la decisión

Los bloques 1 y 2 van con fórmulas, así que se actualizan solos cuando se llenan
los porcentajes del Escenario 3. El bloque 3 se calcula al generar el archivo,
porque en ese momento las columnas de Excel todavía no tienen valores.
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
COL_PUNT = "Puntaje Técnico\n(/5.0)"
COL_FACTOR = "Factor Logistico"
COL_CANT = "Cantidad Esperada"


def _ganadores_por_item(hoja_df, ronda_vigente):
    """
    Por ítem: quién gana en E1 (más barato) y en E2 (mejor puntaje, desempatando
    por precio). Se calcula en Python porque en el Excel esas columnas son
    fórmulas que todavía no tienen valor cuando se genera el archivo.
    """
    col_pu = f"Precio Unitario R{ronda_vigente}\n(USD)"
    if COL_ID not in hoja_df.columns or col_pu not in hoja_df.columns:
        return {}

    df = hoja_df.copy()
    pu = pd.to_numeric(df[col_pu], errors="coerce")
    factor = (pd.to_numeric(df[COL_FACTOR], errors="coerce").fillna(0)
              if COL_FACTOR in df.columns else 0)
    df["_costo"] = pu * (1 + factor)
    df["_cant"] = (pd.to_numeric(df[COL_CANT], errors="coerce").fillna(0)
                   if COL_CANT in df.columns else 1)
    df["_punt"] = (pd.to_numeric(df[COL_PUNT], errors="coerce")
                   if COL_PUNT in df.columns else pd.Series([None] * len(df)))

    salida = {}
    for item_id, grupo in df.groupby(COL_ID):
        validos = grupo[grupo["_costo"].notna() & (grupo["_costo"] > 0)]
        if validos.empty:
            continue
        fila_e1 = validos.loc[validos["_costo"].idxmin()]
        e1 = (fila_e1[COL_PROV], float(fila_e1["_costo"] * fila_e1["_cant"]))

        con_punt = validos[validos["_punt"].notna() & (validos["_punt"] > 0)]
        if con_punt.empty:
            e2 = (None, 0.0)
        else:
            mejor = con_punt["_punt"].max()
            empate = con_punt[con_punt["_punt"] >= mejor - 1e-9]
            fila_e2 = empate.loc[empate["_costo"].idxmin()]
            e2 = (fila_e2[COL_PROV], float(fila_e2["_costo"] * fila_e2["_cant"]))

        salida[item_id] = {
            "E1": e1, "E2": e2,
            "destino": fila_e1.get(COL_DEST),
            "descripcion": fila_e1.get(COL_DESC),
        }
    return salida


def escribir_resumen_ejecutivo(wb, hoja_df, prop_rango, columnas_prop,
                                ronda_actual=1, ronda_vigente=None,
                                nombre_proceso=""):
    fila_ini, fila_fin = prop_rango
    hoja = "'Prop Adj Final'"
    ronda_vigente = ronda_vigente or ronda_actual

    def col(nombre):
        for i, c in enumerate(columnas_prop, start=1):
            if str(c).strip() == str(nombre).strip():
                return get_column_letter(i)
        return None

    def rng(nombre):
        letra = col(nombre)
        return f"{hoja}!${letra}${fila_ini}:${letra}${fila_fin}" if letra else None

    r_dest, r_prov, r_id = rng(COL_DEST), rng(COL_PROV), rng(COL_ID)
    montos = {e: rng(f"Monto {e}\n(USD)") for e in ("E1", "E2", "E3")}
    pcts = {"E1": rng("% Adj. E1"), "E2": rng("% Adj. E2"),
            "E3": rng("% Adj. E3\n← MANUAL") or rng("% Adj. E3")}
    if not r_dest or not r_prov or not montos["E1"] or COL_DEST not in hoja_df.columns:
        return None

    n_cols = 10
    ws = wb.create_sheet("Resumen Ejecutivo", 1)

    ganadores = _ganadores_por_item(hoja_df, ronda_vigente)
    en_disputa = [i for i, g in ganadores.items()
                  if g["E2"][0] is not None and g["E1"][0] != g["E2"][0]]
    total_items = len(ganadores)

    _titulo(ws, f"RESUMEN EJECUTIVO — {nombre_proceso or 'proceso'} · Ronda "
                f"{ronda_actual} (USD)", n_cols)
    _nota(ws, f"De los {total_items} ítems, {total_items - len(en_disputa)} coinciden "
              "entre el más económico (E1) y el de mejor puntaje técnico (E2): se "
              f"pueden aprobar en bloque. Los {len(en_disputa)} del bloque 3 requieren "
              "decisión.", n_cols)

    # ── Bloque 1: monto por destino ──────────────────────────────────────────
    fila = 4
    ws.cell(row=fila, column=1, value="1. MONTO ADJUDICADO POR DESTINO").font = Font(
        name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
    fila += 1
    _headers(ws, ["País destino", "Ítems", "Monto E1", "Monto E2",
                  "Monto E3 (propuesta)", "Proveedor principal en E3",
                  "", "", "", ""], row=fila)

    destinos = [d for d in hoja_df[COL_DEST].dropna().unique()]
    f_dest = fila + 1
    for i, dest in enumerate(destinos):
        r = f_dest + i
        ws.cell(row=r, column=1, value=dest)
        ws.cell(row=r, column=2, value=f'=COUNTIFS({r_dest},$A{r},{pcts["E1"]},">0")')
        for j, e in enumerate(("E1", "E2", "E3")):
            ws.cell(row=r, column=3 + j,
                    value=f'=SUMIFS({montos[e]},{r_dest},$A{r})' if montos[e] else None)
        if montos["E3"]:
            ws.cell(row=r, column=6, value=(
                f'=IF(E{r}=0,"—",IFERROR(INDEX({r_prov},SUMPRODUCT(MATCH('
                f'SUMPRODUCT(MAX(({r_dest}=$A{r})*{montos["E3"]})),'
                f'({r_dest}=$A{r})*{montos["E3"]},0))),"—"))'))
    _estilo_datos(ws, len(destinos), 6, header_row=fila)

    r_tot = f_dest + len(destinos)
    ws.cell(row=r_tot, column=1, value="TOTAL").font = Font(
        name=FONT_NAME, size=9, bold=True)
    for c in range(2, 6):
        letra = get_column_letter(c)
        celda = ws.cell(row=r_tot, column=c,
                        value=f"=SUM({letra}{f_dest}:{letra}{r_tot - 1})")
        celda.font = Font(name=FONT_NAME, size=9, bold=True)
    for c in range(1, 7):
        cc = ws.cell(row=r_tot, column=c)
        cc.border = _thin_border()
        cc.fill = PatternFill(start_color=GRIS, end_color=GRIS, fill_type="solid")
    for r in range(f_dest, r_tot + 1):
        for c in range(3, 6):
            ws.cell(row=r, column=c).number_format = "#,##0.00"

    # ── Bloque 2: proveedor por destino ──────────────────────────────────────
    fila = r_tot + 2
    ws.cell(row=fila, column=1,
            value="2. QUÉ PROVEEDOR QUEDA EN CADA DESTINO").font = Font(
        name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
    fila += 1
    _headers(ws, ["Destino", "Proveedor", "Ítems E1", "Monto E1", "Ítems E2",
                  "Monto E2", "Ítems E3", "Monto E3", "", ""], row=fila)
    r = fila + 1
    filas_sub = []
    for dest in destinos:
        provs_dest = sorted(
            hoja_df.loc[hoja_df[COL_DEST] == dest, COL_PROV].dropna().unique())
        inicio = r
        for prov in provs_dest:
            ws.cell(row=r, column=1, value=dest)
            ws.cell(row=r, column=2, value=prov)
            for j, e in enumerate(("E1", "E2", "E3")):
                c = 3 + j * 2
                if pcts[e]:
                    ws.cell(row=r, column=c, value=(
                        f'=COUNTIFS({r_dest},$A{r},{r_prov},$B{r},{pcts[e]},">0")'))
                if montos[e]:
                    celda = ws.cell(row=r, column=c + 1, value=(
                        f'=SUMIFS({montos[e]},{r_dest},$A{r},{r_prov},$B{r})'))
                    celda.number_format = "#,##0.00"
            r += 1
        ws.cell(row=r, column=1, value=dest)
        ws.cell(row=r, column=2, value="SUBTOTAL").font = Font(
            name=FONT_NAME, size=9, bold=True)
        for c in range(3, 9):
            letra = get_column_letter(c)
            celda = ws.cell(row=r, column=c,
                            value=f"=SUM({letra}{inicio}:{letra}{r - 1})")
            celda.font = Font(name=FONT_NAME, size=9, bold=True)
            celda.number_format = "#,##0.00" if c % 2 == 0 else "#,##0"
        filas_sub.append(r)
        r += 1
    if r - fila - 1 > 0:
        _estilo_datos(ws, r - fila - 1, 8, header_row=fila)
    for rs in filas_sub:
        for c in range(1, 9):
            ws.cell(row=rs, column=c).fill = PatternFill(
                start_color=GRIS, end_color=GRIS, fill_type="solid")

    # ── Bloque 3: ítems a decidir ────────────────────────────────────────────
    fila = r + 1
    ws.cell(row=fila, column=1,
            value="3. ÍTEMS A DECIDIR EN LA REUNIÓN — el más económico y el de "
                  "mejor puntaje técnico no coinciden").font = Font(
        name=FONT_NAME, size=10, bold=True, color=NAVY[2:])
    fila += 1
    _headers(ws, ["ID SAP", "Destino", "Descripción", "E1 — Proveedor", "Monto E1",
                  "E2 — Proveedor", "Monto E2", "E3 — Proveedor (propuesta)",
                  "Monto E3", "Decisión (marcar aquí)"], row=fila)
    f_dec = fila + 1

    for i, item_id in enumerate(en_disputa):
        g = ganadores[item_id]
        r = f_dec + i
        ws.cell(row=r, column=1, value=item_id)
        ws.cell(row=r, column=2, value=g["destino"])
        ws.cell(row=r, column=3, value=g["descripcion"])
        ws.cell(row=r, column=4, value=g["E1"][0])
        ws.cell(row=r, column=5, value=round(g["E1"][1], 2))
        ws.cell(row=r, column=6, value=g["E2"][0])
        ws.cell(row=r, column=7, value=round(g["E2"][1], 2))
        if montos["E3"] and pcts["E3"]:
            ws.cell(row=r, column=8, value=(
                f'=IF(SUMIFS({montos["E3"]},{r_id},$A{r})=0,"",'
                f'IFERROR(INDEX({r_prov},SUMPRODUCT(MATCH(1,'
                f'({r_id}=$A{r})*({pcts["E3"]}>0),0))),""))'))
            ws.cell(row=r, column=9, value=f'=SUMIFS({montos["E3"]},{r_id},$A{r})')
        for c in (5, 7, 9):
            ws.cell(row=r, column=c).number_format = "#,##0.00"

    if en_disputa:
        _estilo_datos(ws, len(en_disputa), n_cols, header_row=fila, input_cols=(10,))
    else:
        ws.cell(row=f_dec, column=1,
                value="Sin ítems en disputa: el más económico y el de mejor puntaje "
                      "técnico coinciden en todos los ítems.")

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 46
    for c in range(4, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 20
    ws.sheet_view.showGridLines = False
    return ws
