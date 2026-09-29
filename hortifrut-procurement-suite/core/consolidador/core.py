"""
Consolidador de Cotizaciones de Proveedores — Hortifrut
Lógica central: parseo, normalización, merge y benchmarking
"""

import pandas as pd
import re
from pathlib import Path
from typing import Optional


# ─── Normalización de headers ────────────────────────────────────────────────

def normalizar_header(nombre: str) -> str:
    """Quita saltos de línea, espacios extra y asteriscos de los nombres de columna."""
    if not isinstance(nombre, str):
        return str(nombre)
    return re.sub(r'\s+', ' ', nombre.replace('\n', ' ').replace('*', '')).strip()


def normalizar_headers(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [normalizar_header(c) for c in df.columns]
    return df


# ─── Detección dinámica de hoja ──────────────────────────────────────────────

KEYWORDS_COTIZACION = ["cotización", "cotizacion", "anexo ii"]

def detectar_hoja_cotizacion(xl: pd.ExcelFile) -> Optional[str]:
    """Encuentra la hoja de cotización aunque el nombre varíe entre versiones."""
    for hoja in xl.sheet_names:
        nombre_lower = hoja.lower()
        if any(kw in nombre_lower for kw in KEYWORDS_COTIZACION):
            return hoja
    return None


KEYWORDS_COND_COMERCIALES = ["condiciones comerciales", "condiciones"]

def detectar_hoja_condiciones(xl: pd.ExcelFile) -> Optional[str]:
    for hoja in xl.sheet_names:
        if any(kw in hoja.lower() for kw in KEYWORDS_COND_COMERCIALES):
            return hoja
    return None


# ─── Extracción de datos del proveedor ───────────────────────────────────────

def extraer_razon_social(xl: pd.ExcelFile) -> str:
    """Lee la razón social desde la hoja Condiciones Comerciales."""
    hoja = detectar_hoja_condiciones(xl)
    if not hoja:
        return "Proveedor Desconocido"
    try:
        df = pd.read_excel(xl, sheet_name=hoja, header=None)
        # Busca la fila con "Razón Social" y toma el valor de la columna 1
        for i, row in df.iterrows():
            celda = str(row.iloc[0]).lower()
            if "razón social" in celda or "razon social" in celda:
                valor = str(row.iloc[1]).strip()
                if valor and valor.lower() not in ("nan", "ingrese razón social completa", ""):
                    return valor
    except Exception:
        pass
    return "Proveedor Desconocido"


# ─── Parser principal de archivo de proveedor ────────────────────────────────

# Mapeo flexible: nombre canónico → posibles variantes en el Excel
COLUMNAS_REQUERIDAS = {
    "ID":                ["id"],
    "Descripcion":       ["descripción", "descripcion"],
    "Pais":              ["país solicitante", "país", "pais solicitante", "pais"],
    "Sociedad":          ["sociedad"],
    "Partida_Principal": ["partida principal", "partida\nprincipal"],
    "Subcategoria":      ["subcategoría", "subcategoria"],
    "Codigo_SAP":        ["código sap", "codigo sap"],
    "Cantidad_Total":    ["cantidad total", "cantidad\ntotal"],
    "UM":                ["um"],
    "Cotiza":            ["¿cotiza? (sí/no)", "¿cotiza?\n(sí/no)", "cotiza"],
    "Marca_Ofertada":    ["marca ofertada"],
    "Modelo_Ofertado":   ["modelo ofertado"],
    "PU_EXW":            ["p.u. exw\n(usd)", "p.u. exw (usd)", "p.u (exw)", "p.u. usd\n(referencial)"],
    "Costo_Flete":       ["costo flete\n(usd)", "costo flete (usd)", "costo de envio"],
    "PU_Total":          ["p.u. total\n(usd)", "p.u. total (usd)"],
    "Subtotal":          ["subtotal\n(usd)", "subtotal (usd)"],
    "Tiempo_Entrega":    ["tiempo de entrega\n(días háb.)", "tiempo entrega\n(días háb.)", "tiempo de entrega"],
    "Garantia":          ["tiempo de garantía", "t. garantía"],
    "Modalidad":         ["modalidad de transporte"],
}


def mapear_columnas(df: pd.DataFrame) -> dict:
    """
    Retorna un dict {nombre_canonico: nombre_real_en_df}.
    Usa matching por contenido normalizado.
    """
    cols_df = {normalizar_header(c).lower(): c for c in df.columns}
    mapa = {}
    for canon, variantes in COLUMNAS_REQUERIDAS.items():
        for variante in variantes:
            variante_norm = normalizar_header(variante).lower()
            if variante_norm in cols_df:
                mapa[canon] = cols_df[variante_norm]
                break
    return mapa


def parse_supplier_file(filepath: str | Path, nombre_proceso: str = "", ronda: int = 1,
                         nombre_original: str | None = None) -> dict:
    """
    Parsea un archivo Excel de proveedor.
    nombre_original: nombre real del archivo subido (ej. "NETAFIM PERÚ.xlsx"), usado
                      como respaldo del nombre del proveedor si no se encuentra la razón
                      social. Si no se pasa, se usa el nombre del archivo en disco (filepath).
    Retorna:
      - df: DataFrame limpio con columnas canónicas
      - proveedor: nombre del proveedor
      - warnings: lista de advertencias
      - errores: lista de errores fatales (si hay errores, df puede estar vacío)
    """
    filepath = Path(filepath)
    nombre_fallback = Path(nombre_original).stem if nombre_original else filepath.stem
    warnings = []
    errores = []

    # ── 1. Abrir archivo ─────────────────────────────────────────────────────
    try:
        xl = pd.ExcelFile(filepath)
    except Exception as e:
        return {"df": pd.DataFrame(), "proveedor": nombre_fallback,
                "warnings": [], "errores": [f"No se pudo abrir el archivo: {e}"]}

    # ── 2. Extraer razón social ──────────────────────────────────────────────
    proveedor = extraer_razon_social(xl)
    if proveedor == "Proveedor Desconocido":
        proveedor = nombre_fallback  # Usa nombre del archivo subido como fallback
        warnings.append("No se encontró razón social en 'Condiciones Comerciales'. Se usó el nombre del archivo.")

    # ── 3. Detectar hoja de cotización ───────────────────────────────────────
    hoja = detectar_hoja_cotizacion(xl)
    if not hoja:
        errores.append(f"No se encontró hoja de cotización. Hojas disponibles: {xl.sheet_names}")
        return {"df": pd.DataFrame(), "proveedor": proveedor, "warnings": warnings, "errores": errores}

    # ── 4. Leer con header en fila 1 (índice 1) ──────────────────────────────
    try:
        df = pd.read_excel(xl, sheet_name=hoja, header=1)
    except Exception as e:
        errores.append(f"Error leyendo hoja '{hoja}': {e}")
        return {"df": pd.DataFrame(), "proveedor": proveedor, "warnings": warnings, "errores": errores}

    if df.empty:
        errores.append(f"La hoja '{hoja}' está vacía.")
        return {"df": pd.DataFrame(), "proveedor": proveedor, "warnings": warnings, "errores": errores}

    # ── 5. Mapear columnas ───────────────────────────────────────────────────
    mapa = mapear_columnas(df)

    columnas_criticas = ["ID", "Cotiza", "PU_Total"]
    for col in columnas_criticas:
        if col not in mapa:
            warnings.append(f"Columna crítica no encontrada: '{col}'. El archivo puede tener un formato inesperado.")

    # ── 6. Renombrar a nombres canónicos ─────────────────────────────────────
    rename_dict = {v: k for k, v in mapa.items()}
    df = df.rename(columns=rename_dict)

    # ── 7. Filtrar filas con ID válido ───────────────────────────────────────
    if "ID" in df.columns:
        df = df[df["ID"].notna() & (df["ID"].astype(str).str.strip() != "")]
        df = df[~df["ID"].astype(str).str.contains("⬇|INFORMACIÓN|ID", na=False)]
    else:
        warnings.append("Columna ID no encontrada. Se mantendrán todas las filas con datos.")
        df = df.dropna(how='all')

    if df.empty:
        errores.append("No quedaron filas de ítems después de filtrar. Verifica el archivo.")
        return {"df": pd.DataFrame(), "proveedor": proveedor, "warnings": warnings, "errores": errores}

    # ── 8. Filtrar solo ítems cotizados ──────────────────────────────────────
    total_items = len(df)
    if "Cotiza" in df.columns:
        marca_si = df["Cotiza"].astype(str).str.strip().str.lower().isin(["sí", "si", "s", "yes", "y", "1"])
        # Incluir también filas donde Cotiza está vacío PERO hay precio ingresado
        tiene_precio = pd.Series([False] * len(df), index=df.index)
        for col_precio in ["PU_Total", "PU_EXW", "Subtotal"]:
            if col_precio in df.columns:
                vals = pd.to_numeric(df[col_precio], errors="coerce")
                tiene_precio = tiene_precio | (vals.notna() & (vals > 0))
        df_cotiza = df[marca_si | tiene_precio]
        # Excluir solo los que explícitamente dicen "no"
        marca_no = df["Cotiza"].astype(str).str.strip().str.lower().isin(["no", "n", "0"])
        df_cotiza = df_cotiza[~marca_no]
        no_cotiza = total_items - len(df_cotiza)
        if no_cotiza > 0:
            warnings.append(f"{no_cotiza} ítem(s) marcados como 'No cotiza' o sin datos fueron excluidos.")
        df = df_cotiza
    else:
        warnings.append("No se encontró columna '¿Cotiza?'. Se incluirán todos los ítems.")

    if df.empty:
        warnings.append("Este proveedor no cotizó ningún ítem.")
        return {"df": pd.DataFrame(), "proveedor": proveedor, "warnings": warnings, "errores": errores}

    # ── 9. Convertir columnas numéricas ──────────────────────────────────────
    cols_numericas = ["PU_EXW", "Costo_Flete", "PU_Total", "Subtotal", "Cantidad_Total", "Tiempo_Entrega"]
    for col in cols_numericas:
        if col in df.columns:
            original = df[col].copy()
            df[col] = pd.to_numeric(df[col], errors='coerce')
            fallidos = original.notna() & df[col].isna()
            if fallidos.sum() > 0:
                warnings.append(f"Columna '{col}': {fallidos.sum()} valor(es) no numérico(s) convertidos a vacío.")

    # ── 10. PU_Total: recalcular si no existe pero sí EXW y Flete ────────────
    if "PU_Total" not in df.columns or df["PU_Total"].isna().all():
        if "PU_EXW" in df.columns and "Costo_Flete" in df.columns:
            df["PU_Total"] = df["PU_EXW"].fillna(0) + df["Costo_Flete"].fillna(0)
            warnings.append("Columna 'PU_Total' recalculada como EXW + Flete.")
        elif "PU_EXW" in df.columns:
            df["PU_Total"] = df["PU_EXW"]
            warnings.append("Columna 'PU_Total' tomada de EXW (sin flete disponible).")

    # ── 11. Añadir metadatos ─────────────────────────────────────────────────
    df["_proveedor"] = proveedor
    df["_archivo"] = filepath.name
    df["_ronda"] = ronda
    df["_proceso"] = nombre_proceso

    # Solo conservar columnas canónicas conocidas + metadatos
    cols_output = [c for c in (list(COLUMNAS_REQUERIDAS.keys()) + ["_proveedor", "_archivo", "_ronda", "_proceso"])
                   if c in df.columns]
    df = df[cols_output].reset_index(drop=True)

    return {"df": df, "proveedor": proveedor, "warnings": warnings, "errores": errores}


# ─── Consolidador ────────────────────────────────────────────────────────────

def consolidar_proveedores(resultados: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Recibe lista de resultados de parse_supplier_file.
    Retorna:
      - df_consolidado: tabla wide con columnas por proveedor
      - df_alertas: log de ítems con problemas
    """
    dfs = []
    for r in resultados:
        if not r["df"].empty:
            dfs.append(r["df"])

    if not dfs:
        return pd.DataFrame(), pd.DataFrame()

    df_long = pd.concat(dfs, ignore_index=True)

    if "ID" not in df_long.columns:
        return df_long, pd.DataFrame()

    df_long["ID"] = df_long["ID"].astype(str).str.strip()

    # ── Tabla wide: una fila por ID, columnas por proveedor ──────────────────
    proveedores = df_long["_proveedor"].unique()

    # Columnas de referencia (tomadas del primer proveedor que tenga datos)
    cols_ref = ["ID", "Descripcion", "Pais", "Sociedad", "Partida_Principal",
                "Subcategoria", "Codigo_SAP", "Cantidad_Total", "UM"]
    cols_ref = [c for c in cols_ref if c in df_long.columns]

    df_ref = df_long.drop_duplicates(subset=["ID"])[cols_ref].set_index("ID")

    # Columnas de precio por proveedor
    cols_precio = ["PU_EXW", "Costo_Flete", "PU_Total", "Subtotal",
                   "Marca_Ofertada", "Tiempo_Entrega", "Garantia", "Modalidad"]

    frames_prov = []
    for prov in proveedores:
        df_p = df_long[df_long["_proveedor"] == prov].set_index("ID")
        cols_disponibles = [c for c in cols_precio if c in df_p.columns]
        df_p = df_p[cols_disponibles]
        # Abreviar nombre proveedor para header
        tag = prov[:20].replace(" ", "_")
        df_p.columns = [f"{c}__{tag}" for c in df_p.columns]
        frames_prov.append(df_p)

    df_wide = df_ref.copy()
    for fp in frames_prov:
        df_wide = df_wide.join(fp, how="left")

    # ── Calcular benchmarking ─────────────────────────────────────────────────
    cols_pu_total = [c for c in df_wide.columns if c.startswith("PU_Total__")]
    cols_pu_exw = [c for c in df_wide.columns if c.startswith("PU_EXW__")]

    if len(cols_pu_total) >= 1:
        df_wide["PRECIO_MIN"] = df_wide[cols_pu_total].min(axis=1)
        df_wide["PRECIO_MAX"] = df_wide[cols_pu_total].max(axis=1)
        if len(cols_pu_total) >= 2:
            df_wide["DELTA_%"] = ((df_wide["PRECIO_MAX"] - df_wide["PRECIO_MIN"]) /
                                   df_wide["PRECIO_MIN"].replace(0, float('nan')) * 100).round(1)
            # Proveedor ganador (menor precio total, incluye flete)
            df_wide["PROVEEDOR_GANADOR"] = df_wide[cols_pu_total].idxmin(axis=1).str.replace("PU_Total__", "").str.replace("_", " ")

    # Mínimo de EXW por ítem — base exclusiva para el semáforo (Excel y Word)
    if len(cols_pu_exw) >= 1:
        df_wide["PRECIO_MIN_EXW"] = df_wide[cols_pu_exw].min(axis=1)

    df_wide = df_wide.reset_index()

    # ── Log de alertas ───────────────────────────────────────────────────────
    alertas = []
    for r in resultados:
        for w in r["warnings"]:
            alertas.append({"Proveedor": r["proveedor"], "Tipo": "Advertencia", "Detalle": w})
        for e in r["errores"]:
            alertas.append({"Proveedor": r["proveedor"], "Tipo": "Error", "Detalle": e})

    # Ítems sin cotización de algún proveedor
    for col in cols_pu_total:
        tag = col.replace("PU_Total__", "").replace("_", " ")
        sin_precio = df_wide[df_wide[col].isna()]["ID"].tolist()
        for item_id in sin_precio:
            alertas.append({"Proveedor": tag, "Tipo": "Sin cotización", "Detalle": f"Ítem {item_id} no fue cotizado"})

    df_alertas = pd.DataFrame(alertas) if alertas else pd.DataFrame(columns=["Proveedor", "Tipo", "Detalle"])

    return df_wide, df_alertas


def preparar_items_para_word(df_wide: pd.DataFrame, proveedor: str, tag: str) -> list[dict]:
    """
    Extrae y calcula el feedback (semáforo) por ítem para un proveedor específico,
    en el formato que espera el generador de Word.
    """
    col_pu = f"PU_Total__{tag}"
    col_exw = f"PU_EXW__{tag}"
    col_marca = f"Marca_Ofertada__{tag}"

    if col_pu not in df_wide.columns:
        return []

    items = []
    for _, row in df_wide.iterrows():
        pu_total = row[col_pu]
        if pd.isna(pu_total):
            continue

        pu_exw_val = row.get(col_exw)
        if pu_exw_val is None or pd.isna(pu_exw_val):
            pu_exw_val = float(pu_total) * 0.95  # estimado si no viene desglosado

        # Semáforo calculado en base al P.U. EXW (no al total con flete)
        exw_min = row.get("PRECIO_MIN_EXW", pu_exw_val)
        delta_item = ((float(pu_exw_val) - exw_min) / exw_min * 100) if exw_min and exw_min > 0 else 0

        if delta_item <= 5:
            feedback = "Azul"
        elif delta_item <= 25:
            feedback = "Verde"
        elif delta_item <= 50:
            feedback = "Amarillo"
        else:
            feedback = "Rojo"

        marca_val = row.get(col_marca)
        if marca_val is None or pd.isna(marca_val):
            marca_val = ""

        items.append({
            "id": str(row.get("ID", "")),
            "descripcion": str(row.get("Descripcion", ""))[:80],
            "marca_ofertada": str(marca_val)[:60],
            "pu_exw": round(float(pu_exw_val), 2),
            "costo_flete": round(float(pu_total) - float(pu_exw_val), 2),
            "pu_total": round(float(pu_total), 2),
            "feedback": feedback,
        })

    return items


def generar_word_feedback(df_wide: pd.DataFrame, proveedor: str, nombre_proceso: str,
                           ronda: int) -> tuple[bytes | None, str]:
    """
    Genera el Word de feedback económico para un proveedor.
    Antes se delegaba a un script de Node.js (`generar_feedback_word.js`); ahora se
    construye con python-docx en `word_feedback.py`, sin dependencias externas.
    Retorna (bytes_del_docx_o_None, mensaje_de_error_o_vacio).
    """
    from .word_feedback import generar_docx

    tag = proveedor[:20].replace(" ", "_")
    items = preparar_items_para_word(df_wide, proveedor, tag)
    if not items:
        return None, "No hay ítems con precio para este proveedor."

    try:
        return generar_docx(nombre_proceso, ronda, proveedor, items), ""
    except Exception as e:
        return None, f"No se pudo generar el documento Word: {e}"


def exportar_consolidado(df_wide: pd.DataFrame, df_alertas: pd.DataFrame,
                          nombre_proceso: str, ronda: int) -> bytes:
    """Genera el archivo Excel consolidado en memoria (bytes) para descarga."""
    import io
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_wide.to_excel(writer, sheet_name="Consolidado", index=False)
        # Hoja Alertas desactivada

        wb = writer.book
        ws = writer.sheets["Consolidado"]

        # Colores
        VERDE       = "C6EFCE"
        AMARILLO    = "FFEB9C"
        ROJO        = "FFC7CE"
        AZUL        = "BDD7EE"  # azul claro para semáforo (distinto del azul del header)
        AZUL_HEADER = "1F4E79"
        GRIS_REF    = "D9E1F2"
        NARANJA_BM  = "FCE4D6"

        thin = Side(style='thin', color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        cols = list(df_wide.columns)

        # Clasificar columnas
        cols_referencia  = [c for c in cols if not any(sep in c for sep in ["__", "PRECIO", "DELTA", "PROVEEDOR_GANADOR"])]
        cols_proveedor   = [c for c in cols if "__" in c]
        cols_benchmark   = [c for c in cols if c in ["PRECIO_MIN", "PRECIO_MAX", "DELTA_%", "PROVEEDOR_GANADOR"]]

        # ── Dar formato al header ─────────────────────────────────────────────
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

            if col_name in cols_referencia:
                cell.fill = PatternFill("solid", fgColor=AZUL_HEADER)
            elif col_name in cols_proveedor:
                cell.fill = PatternFill("solid", fgColor="2E75B6")
            elif col_name in cols_benchmark:
                cell.fill = PatternFill("solid", fgColor="833C00" if "GANADOR" not in col_name else "375623")

        ws.row_dimensions[1].height = 35

        # ── Dar formato a las celdas de datos ─────────────────────────────────
        cols_pu_exw = [i+1 for i, c in enumerate(cols) if c.startswith("PU_EXW__")]

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center")
                col_name = cols[cell.column - 1]

                # Fondo de columnas de referencia
                if col_name in cols_referencia:
                    cell.fill = PatternFill("solid", fgColor=GRIS_REF)

                # Fondo de columnas benchmark
                if col_name in ["PRECIO_MIN", "PRECIO_MAX", "DELTA_%", "PROVEEDOR_GANADOR"]:
                    cell.fill = PatternFill("solid", fgColor=NARANJA_BM)

                # Formato numérico
                if col_name in ["PU_EXW", "Costo_Flete", "PU_Total", "Subtotal",
                                 "PRECIO_MIN", "PRECIO_MAX"] or "__" in col_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

                if col_name == "DELTA_%":
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.0"%"'

            # ── Semaforización en columnas PU_EXW por proveedor (base: EXW, no total) ──
            if len(cols_pu_exw) >= 2:
                exw_vals = [row[i-1].value for i in cols_pu_exw if isinstance(row[i-1].value, (int, float))]
                if exw_vals:
                    exw_min = min(exw_vals)
                    for col_idx in cols_pu_exw:
                        c = row[col_idx - 1]
                        if isinstance(c.value, (int, float)) and exw_min > 0:
                            delta_item = (c.value - exw_min) / exw_min * 100
                            if delta_item <= 5:
                                c.fill = PatternFill("solid", fgColor=AZUL)
                            elif delta_item <= 25:
                                c.fill = PatternFill("solid", fgColor=VERDE)
                            elif delta_item <= 50:
                                c.fill = PatternFill("solid", fgColor=AMARILLO)
                            else:
                                c.fill = PatternFill("solid", fgColor=ROJO)

        # ── Anchos de columna ─────────────────────────────────────────────────
        for col_idx, col_name in enumerate(cols, 1):
            col_letter = get_column_letter(col_idx)
            if col_name == "Descripcion":
                ws.column_dimensions[col_letter].width = 35
            elif col_name in cols_benchmark:
                ws.column_dimensions[col_letter].width = 18
            elif "__" in col_name:
                ws.column_dimensions[col_letter].width = 16
            else:
                ws.column_dimensions[col_letter].width = 14

        ws.freeze_panes = "D2"

        # Formato hoja Alertas desactivado

    output.seek(0)
    return output.read()
