"""
Estilo visual para la hoja 'Prop Adj Final', replicando el look & feel de la
plantilla de referencia de Hortifrut:
  - Fila 1: banner de título (fondo azul marino, texto blanco)
  - Fila 2: panel de parámetros (Moneda asumida, umbrales Verde ≤ / Amarillo ≤)
            + banners de Escenario 1/2/3 (fondo azul medio)
  - Fila 3: encabezados de columna (fondo azul marino; azul medio para las
            3 columnas de costo original EXW/Envío/Adicionales)
  - Datos desde fila 4
  - Semáforos con emoji (🟢/🟡/🔴/S/BASE) calculados por fórmula + formato
    condicional real (mismo color de fondo que el texto indica)
  - Formatos numéricos: moneda "$#,##0.00", porcentaje "0.0%"
  - Freeze panes en C4
"""

import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string

NAVY = "FF1F4E79"
MEDIUM_BLUE = "FF2F75B6"
ESCENARIO_BLUE = "FF2F5597"
PARAM_BG = "FFD6DCE4"
FACTOR_BG = "FFFFF2CC"
VERDE = "FFC6EFCE"
AMARILLO = "FFFFEB9C"
ROJO = "FFFFC7CE"
GRIS = "FFE7E6E6"
WHITE = "FFFFFFFF"
FONT_NAME = "Calibri"

HEADER_ROW = 3
DATA_START_ROW = 4

GREEN_HEADER = "FF00B050"

# Columnas nuevas que van con encabezado verde (agregadas respecto de la
# plantilla original de Coupa).
GREEN_HEADERS = (
    "país destino", "incoterm", "factor logistico", "factor logístico",
    "precio histórico 2526 + factor", "indice de inflación", "índice de inflación",
)

# Bloque de precios por ronda (encabezado azul medio).
EXW_BLOCK_HINTS = ("precio unitario r", "costo total  (usd) r", "costo total (usd) r")

HOMOLOGACION_OK = ["HOMOLOGADO", "APROBADO"]
HOMOLOGACION_WARN = ["POR REVISAR"]
HOMOLOGACION_BAD = ["NO HOMOLOGADO", "OBSOLETO", "RECHAZADO"]
HOMOLOGACION_NA = ["S/DATOS"]


def _thin_border():
    thin = Side(style="thin", color="FFB7B7B7")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _norm(s):
    return re.sub(r"\s+", " ", str(s).replace("\n", " ")).strip().lower()


def _is_money_col(name):
    n = _norm(name)
    if "incoterm" in n.replace(" ", ""):
        return False
    if "factor" in n or "inflaci" in n:
        return False  # son porcentajes/multiplicadores, no montos
    return ("costo" in n or "precio" in n or "gasto" in n or "ahorro" in n
            or "monto" in n or "baseline" in n or "subtotal" in n) and "%" not in n


def _is_factor_col(name):
    n = _norm(name)
    return "factor logistico" in n or "factor logístico" in n or "inflaci" in n


def _is_pct_col(name):
    n = _norm(name)
    return ("%" in name) or n.startswith("desviación") or n.startswith("vs.") or n.startswith("vs ")


def _find_col(df, header_text):
    return list(df.columns).index(header_text) + 1 if header_text in df.columns else None


def _find_block(df, pattern):
    """Devuelve (primera_col, ultima_col) 1-based de columnas cuyo header matchea el patrón regex.
    El patrón se evalúa contra el header normalizado (guiones bajos incluidos), para que
    columnas como 'Ahorro_E1' también cuenten como parte del bloque del Escenario 1."""
    idxs = [i + 1 for i, c in enumerate(df.columns) if re.search(pattern, _norm(c))]
    if not idxs:
        return None
    return min(idxs), max(idxs)


def escribir_encabezado_y_estilo(ws, df, titulo="PROPUESTA ADJUDICACIÓN FINAL", verde_umbral=0.05, amarillo_umbral=0.10):
    """
    Escribe título (fila1), panel de parámetros + banners de escenario (fila2)
    y encabezados de columna con color por sección (fila3). Devuelve
    {'verde_cell':..., 'amarillo_cell':...} con las direcciones de los
    umbrales, para usar en las fórmulas de semáforo.
    """
    n_cols = len(df.columns)

    # --- Fila 1: título ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1, value=titulo)
    title_cell.font = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)
    title_cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, n_cols + 1):
        ws.cell(row=1, column=c).border = _thin_border()
        ws.cell(row=1, column=c).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws.row_dimensions[1].height = 27

    # --- Fila 2: panel de parámetros ---
    param_cells = [("D", "Moneda asumida", "E", "USD", "General"),
                   ("F", "Verde ≤", "G", verde_umbral, "0%"),
                   ("H", "Amarillo ≤", "I", amarillo_umbral, "0%")]
    for label_col, label, val_col, val, fmt in param_cells:
        lc = ws[f"{label_col}2"]
        lc.value = label
        vc = ws[f"{val_col}2"]
        vc.value = val
        vc.number_format = fmt
        for cell in (lc, vc):
            cell.font = Font(name=FONT_NAME, size=8, bold=True)
            cell.fill = PatternFill(start_color=PARAM_BG, end_color=PARAM_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = _thin_border()

    # --- Banners de Escenario 1/2/3 (fila 2, sobre cada bloque E1/E2/E3) ---
    escenario_labels = {
        1: "ESCENARIO 1 — MÁS ECONÓMICO (precio mínimo entre todos los proveedores)",
        2: "ESCENARIO 2 — MEJOR PUNTAJE TÉCNICO (mayor puntaje del ítem, aunque sea bajo)",
        3: "ESCENARIO 3 — PROPUESTA MANUAL  ← completar % a mano",
    }
    for n, label in escenario_labels.items():
        # Límite "no alfanumérico" en vez de \b, para que 'Ahorro_E1' (guion
        # bajo antes de E1) también cuente como parte del bloque.
        block = _find_block(df, rf"(?<![a-z0-9])e{n}(?![a-z0-9])")
        if not block:
            continue
        c1, c2 = block
        if c2 > c1:
            ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        cell = ws.cell(row=2, column=c1, value=label)
        cell.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=ESCENARIO_BLUE, end_color=ESCENARIO_BLUE, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c in range(c1, c2 + 1):
            ws.cell(row=2, column=c).border = _thin_border()
            ws.cell(row=2, column=c).fill = PatternFill(start_color=ESCENARIO_BLUE, end_color=ESCENARIO_BLUE, fill_type="solid")
    ws.row_dimensions[2].height = 22

    # --- Fila 3: encabezados de columna ---
    for c, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=3, column=c, value=col_name)
        normed = _norm(col_name)
        if normed in GREEN_HEADERS:
            color = GREEN_HEADER
        elif any(h in normed for h in EXW_BLOCK_HINTS):
            color = MEDIUM_BLUE
        else:
            color = NAVY
        cell.font = Font(name=FONT_NAME, size=8, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
        max_len = max([len(str(col_name).replace("\n", " "))]
                      + [len(str(v)) for v in df[col_name].astype(str).head(200)])
        ws.column_dimensions[get_column_letter(c)].width = min(max(max_len * 0.9 + 2, 9), 30)
    ws.row_dimensions[3].height = 40

    ws.freeze_panes = "C4"
    return {"verde_cell": "$G$2", "amarillo_cell": "$I$2"}


def formatear_datos(ws, df, header_row=HEADER_ROW, data_start_row=DATA_START_ROW):
    """Aplica fuente, alineación, bordes y formato numérico a las filas de datos."""
    n_rows = len(df)
    for c, col_name in enumerate(df.columns, start=1):
        money = _is_money_col(col_name)
        pct = _is_pct_col(col_name)
        factor = _is_factor_col(col_name)
        left = _norm(col_name).startswith("descripci")
        for r in range(data_start_row, data_start_row + n_rows):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT_NAME, size=9, color="FF000000")
            cell.alignment = Alignment(horizontal="left" if left else "center",
                                        vertical="center", wrap_text=True)
            cell.border = _thin_border()
            if factor:
                cell.number_format = "0.00%"
            elif pct:
                cell.number_format = "0.0%"
            elif money:
                cell.number_format = "$#,##0.00"


def aplicar_formato_condicional(ws, df, header_row=HEADER_ROW, data_start_row=DATA_START_ROW,
                                 verde_cell="$G$2", amarillo_cell="$I$2"):
    """
    Formato condicional:
      - 'Calidad / Homologación': verde/amarillo/rojo/gris por texto.
      - 'Desviación con Mín [%]': verde si ≤ umbral verde, amarillo si ≤ umbral
        amarillo, rojo si lo supera (usa los umbrales editables de la fila 2).
      - 'Ahorro_E{n}': rojo si el ahorro es negativo (el precio subió respecto
        del Baseline), para que salte a la vista.
    """
    n_rows = len(df)
    last_row = data_start_row + n_rows - 1
    if n_rows == 0:
        return

    def _add_text_rules(col_letter, ok_values, warn_values, bad_values, na_values):
        rng = f"{col_letter}{data_start_row}:{col_letter}{last_row}"
        for v in ok_values:
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{v}"'],
                                 fill=PatternFill(start_color=VERDE, end_color=VERDE, fill_type="solid")))
        for v in warn_values:
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{v}"'],
                                 fill=PatternFill(start_color=AMARILLO, end_color=AMARILLO, fill_type="solid")))
        for v in bad_values:
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{v}"'],
                                 fill=PatternFill(start_color=ROJO, end_color=ROJO, fill_type="solid")))
        for v in na_values:
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{v}"'],
                                 fill=PatternFill(start_color=GRIS, end_color=GRIS, fill_type="solid")))

    for i, col_name in enumerate(df.columns, start=1):
        n = _norm(col_name)
        col_letter = get_column_letter(i)
        rng = f"{col_letter}{data_start_row}:{col_letter}{last_row}"

        if "homologaci" in n or "calidad" in n:
            _add_text_rules(col_letter, HOMOLOGACION_OK, HOMOLOGACION_WARN,
                             HOMOLOGACION_BAD, HOMOLOGACION_NA)
        elif n.startswith("desviación") or n.startswith("desviacion"):
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'AND(ISNUMBER({col_letter}{data_start_row}),'
                          f'{col_letter}{data_start_row}<={verde_cell})'],
                fill=PatternFill(start_color=VERDE, end_color=VERDE, fill_type="solid"), stopIfTrue=True))
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'AND(ISNUMBER({col_letter}{data_start_row}),'
                          f'{col_letter}{data_start_row}<={amarillo_cell})'],
                fill=PatternFill(start_color=AMARILLO, end_color=AMARILLO, fill_type="solid"), stopIfTrue=True))
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'ISNUMBER({col_letter}{data_start_row})'],
                fill=PatternFill(start_color=ROJO, end_color=ROJO, fill_type="solid")))
        elif n.startswith("ahorro"):
            ws.conditional_formatting.add(rng, CellIsRule(
                operator="lessThan", formula=["0"],
                fill=PatternFill(start_color=ROJO, end_color=ROJO, fill_type="solid")))
