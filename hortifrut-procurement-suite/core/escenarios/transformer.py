"""
Transforma la hoja 'Prop Adj Final' de un evento de Coupa (formato largo:
una fila por combinación Item + Zona/Lote + Proveedor) en la estructura
'Consolidado' (formato ancho: una fila por Item + Zona/Lote, con columnas
repetidas por cada proveedor que cotizó).

La estructura de salida replica lo más fielmente posible el Consolidado de
referencia:

    ID, Descripcion, Pais, Sociedad, Partida_Principal, Subcategoria,
    Codigo_SAP, Cantidad_Total, UM, Zona_Lote,
    PU_EXW__<Proveedor>, Costo_Flete__<Proveedor>, PU_Total__<Proveedor>,
    Subtotal__<Proveedor>, <Campo_Detectado>__<Proveedor>, ...
    PRECIO_MIN, PRECIO_MAX, DELTA_%, PROVEEDOR_GANADOR, PRECIO_MIN_EXW

Es AGNÓSTICO al número de proveedores y a los campos descriptivos
específicos de cada evento (Marca, Garantía, Homologación, INCOTERM, etc.)
— estos se detectan automáticamente y se nombran con la misma convención
(Palabra_Palabra) que usa la plantilla original.
"""

import re
import unicodedata
import pandas as pd
import openpyxl


# ---------------------------------------------------------------------------
# Columnas base del Consolidado de referencia. Las que no existan en
# 'Prop Adj Final' se generan vacías, para no romper la estructura.
# ---------------------------------------------------------------------------
BASE_COLUMNS = [
    "ID", "Descripcion", "Pais", "Sociedad", "Partida_Principal",
    "Subcategoria", "Codigo_SAP", "Cantidad_Total", "UM", "Zona_Lote",
]

# Orden fijo de los campos de precio por proveedor (idéntico al original).
PRICE_FIELDS_ORDER = ["PU_EXW", "Costo_Flete", "PU_Total", "Subtotal"]


def _norm(s):
    """Normaliza un encabezado para COMPARARLO: sin tildes, sin saltos de línea, minúsculas."""
    if s is None:
        return ""
    s = str(s).replace("\n", " ")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


CORE_PATTERNS = {
    "lote_id_sap": r"^lote \+ id sap$",
    "lote_id_sap_prov": r"^lote \+ id sap \+ proveedor$",
    "id_sap": r"^id sap$",
    "descripcion": r"^descripcion del articulo",
    "cantidad": r"^cantidad esperada$",
    "proveedor": r"^nombre del proveedor$",
    "zona_lote": r"^zona / lote$",
    "costo_exw_orig": r"^costo caja / clamshell",
    "costo_flete_orig": r"^costo de envio",
    "costo_adicional": r"^costos adicionales",
    "costo_total_usd": r"^costo unitario total \(usd\)$",
    "costo_total_ddp": r"^costo unitario total \(ddp\)$",
}

# Columnas de análisis / histórico / escenarios que NO deben pasar como
# columnas dinámicas por proveedor (son cálculos del motor, no atributos
# de la oferta). Se excluyen por patrón para que funcione en cualquier evento.
EXCLUDE_PATTERNS = [
    r"^minimo costo",
    r"^desviacion",
    r"^precio ppto",
    r"^precio referencia",
    r"^precio historico",
    r"^\(%\) adjudicacion ep",
    r"^volumen ep",
    r"^gasto ep",
    r"^gasto ppto",
    r"^gasto t\d",
    r"^vs\. ",
    r"^semaforo",
    r"^costo orig\. princip",
    r"^costo orig\. envio",
    r"^costo ajust\. por incoterm",
    r"^% adj\. e\d",
    r"^vol\. e\d",
    r"^gasto e\d",
    r"^vs ppto% e\d",
    r"^semaforo e\d",
    r"^semaforo\se\d",
]

# Sufijos de tipo de dato que Coupa agrega a los encabezados y que no
# queremos que queden en el nombre de campo final (ej. "INCOTERM (Texto)" -> "Incoterm").
_TYPE_SUFFIX_RE = re.compile(
    r"\s*\((texto|numero|número|entero|fecha)\)\s*$", re.IGNORECASE
)


def _find_header_row(ws, max_scan=10):
    """Busca la fila de encabezados (contiene 'ID SAP' y 'Nombre del proveedor')."""
    for r in range(1, max_scan + 1):
        values = [_norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if "id sap" in values and "nombre del proveedor" in values:
            return r
    raise ValueError("No se encontró la fila de encabezados en 'Prop Adj Final' "
                      "(se esperaba 'ID SAP' y 'Nombre del proveedor').")


def load_prop_adj_final(file_path_or_buffer, sheet_name="Prop Adj Final"):
    """Carga la hoja como DataFrame usando valores calculados (cache de fórmulas)."""
    wb = openpyxl.load_workbook(file_path_or_buffer, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"El archivo no contiene una hoja llamada '{sheet_name}'. "
                          f"Hojas disponibles: {wb.sheetnames}")
    ws = wb[sheet_name]
    header_row = _find_header_row(ws)

    headers = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    headers = [h if h is not None else f"col_{i}" for i, h in enumerate(headers)]

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None for v in values):
            continue
        rows.append(values)

    df = pd.DataFrame(rows, columns=headers)
    return df


def _match_col(df, pattern):
    """Devuelve el primer nombre de columna real cuyo header normalizado matchea el patrón."""
    for col in df.columns:
        if re.search(pattern, _norm(col)):
            return col
    return None


def _clean_provider_name(name):
    """Limpia el nombre del proveedor tal como aparece en Coupa (quita el código SAP inicial)."""
    if name is None:
        return None
    name = str(name)
    name = re.sub(r"^\d+\s+", "", name).strip()  # quita "10023042  " al inicio
    name = re.sub(r"\s+", " ", name)
    return name


def _provider_slug(name, used):
    """
    Convierte el nombre del proveedor en un sufijo de columna, preservando
    tildes/mayúsculas igual que el Consolidado original
    (ej. 'RIEGOS Y TECNOLOGÍA ' -> 'RIEGOS_Y_TECNOLOGÍA_').
    """
    slug = re.sub(r"[^\w]+", "_", name, flags=re.UNICODE)
    slug = slug.strip("_") or "PROVEEDOR"
    base = slug
    i = 2
    while slug in used:
        slug = f"{base}_{i}"
        i += 1
    used.add(slug)
    return slug


def _field_name(header):
    """
    Convierte un encabezado crudo de Coupa en un nombre de campo con la misma
    convención del Consolidado original (Palabra_Palabra, sin tildes, sin
    sufijos de tipo de dato).
    Ej.: 'Calidad /\\nHomologación' -> 'Calidad_Homologacion'
         'INCOTERM\\n(Texto)'      -> 'Incoterm'
         'Moneda\\noriginal'       -> 'Moneda_Original'
    """
    h = str(header).replace("\n", " ")
    h = _TYPE_SUFFIX_RE.sub("", h)
    h = unicodedata.normalize("NFKD", h).encode("ascii", "ignore").decode("ascii")
    words = re.findall(r"[A-Za-z0-9]+", h)
    return "_".join(w.capitalize() for w in words) or "Campo"


def build_consolidado(df, subtotal_mode="pu_total_x_cantidad"):
    """
    Construye el DataFrame ancho tipo 'Consolidado' a partir del DataFrame largo
    de 'Prop Adj Final', replicando la estructura del Consolidado de referencia.
    """
    id_col = _match_col(df, CORE_PATTERNS["id_sap"])
    desc_col = _match_col(df, CORE_PATTERNS["descripcion"])
    qty_col = _match_col(df, CORE_PATTERNS["cantidad"])
    prov_col = _match_col(df, CORE_PATTERNS["proveedor"])
    zona_col = _match_col(df, CORE_PATTERNS["zona_lote"])
    exw_col = _match_col(df, CORE_PATTERNS["costo_exw_orig"])
    flete_col = _match_col(df, CORE_PATTERNS["costo_flete_orig"])
    adicional_col = _match_col(df, CORE_PATTERNS["costo_adicional"])
    total_ddp_col = _match_col(df, CORE_PATTERNS["costo_total_ddp"])
    total_usd_col = _match_col(df, CORE_PATTERNS["costo_total_usd"])

    missing = [n for n, v in [("ID SAP", id_col), ("Descripción", desc_col),
                               ("Cantidad Esperada", qty_col), ("Nombre del proveedor", prov_col)]
               if v is None]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en 'Prop Adj Final': {missing}")

    pu_total_col = total_ddp_col or total_usd_col
    if pu_total_col is None:
        raise ValueError("No se encontró columna de 'Costo Unitario Total' (DDP o USD).")

    # --- columnas dinámicas: todo lo que no sea core ni esté excluido ---
    core_used = {id_col, desc_col, qty_col, prov_col, zona_col, exw_col,
                 flete_col, adicional_col, total_ddp_col, total_usd_col}
    core_used.discard(None)

    lote_id_col = _match_col(df, CORE_PATTERNS["lote_id_sap"])
    lote_id_prov_col = _match_col(df, CORE_PATTERNS["lote_id_sap_prov"])
    core_used.update([c for c in [lote_id_col, lote_id_prov_col] if c])

    dynamic_cols = []
    for col in df.columns:
        if col in core_used:
            continue
        normed = _norm(col)
        if any(re.search(pat, normed) for pat in EXCLUDE_PATTERNS):
            continue
        dynamic_cols.append(col)

    dynamic_field_names = [_field_name(c) for c in dynamic_cols]

    # --- clave de item: ID SAP + Zona/Lote (si existe) ---
    # Descarta filas sin ID SAP o sin proveedor (notas al pie, filas de comentarios, etc.)
    work = df.dropna(subset=[id_col, prov_col]).copy()
    if zona_col:
        work["_ITEM_KEY"] = work[id_col].astype(str) + " | " + work[zona_col].astype(str)
    else:
        work["_ITEM_KEY"] = work[id_col].astype(str)

    work["_PROVEEDOR_LIMPIO"] = work[prov_col].apply(_clean_provider_name)

    # Coupa registra "sin oferta real" como costo 0 (no vacío). Igual que la fórmula
    # original (MINIFS con ">0"), tratamos 0 y negativos como ausencia de dato.
    work["_PU_TOTAL"] = pd.to_numeric(work[pu_total_col], errors="coerce")
    work.loc[work["_PU_TOTAL"] <= 0, "_PU_TOTAL"] = None
    work["_PU_EXW"] = pd.to_numeric(work[exw_col], errors="coerce") if exw_col else None
    if exw_col:
        work.loc[work["_PU_EXW"] <= 0, "_PU_EXW"] = None
    if flete_col and adicional_col:
        work["_COSTO_FLETE"] = pd.to_numeric(work[flete_col], errors="coerce").fillna(0) + \
                                pd.to_numeric(work[adicional_col], errors="coerce").fillna(0)
    elif flete_col:
        work["_COSTO_FLETE"] = pd.to_numeric(work[flete_col], errors="coerce")
    else:
        work["_COSTO_FLETE"] = None

    # --- metadata de item (una fila por _ITEM_KEY), con la estructura BASE_COLUMNS ---
    item_meta_map = {"_ITEM_KEY": "ITEM_KEY", id_col: "ID", desc_col: "Descripcion",
                      qty_col: "Cantidad_Total"}
    if zona_col:
        item_meta_map[zona_col] = "Zona_Lote"

    item_meta = (work[list(item_meta_map.keys())]
                 .rename(columns=item_meta_map)
                 .drop_duplicates(subset=["ITEM_KEY"])
                 .set_index("ITEM_KEY"))

    # columnas base que no existen en Prop Adj Final -> vacías, para mantener
    # la misma forma que el Consolidado de referencia.
    for col in BASE_COLUMNS:
        if col not in item_meta.columns:
            item_meta[col] = None
    item_meta = item_meta[BASE_COLUMNS]

    providers = sorted(work["_PROVEEDOR_LIMPIO"].dropna().unique().tolist())
    used_slugs = set()
    provider_slug = {p: _provider_slug(p, used_slugs) for p in providers}

    result = item_meta.copy()

    for prov in providers:
        slug = provider_slug[prov]
        sub = work[work["_PROVEEDOR_LIMPIO"] == prov].drop_duplicates(subset=["_ITEM_KEY"])
        sub = sub.set_index("_ITEM_KEY")

        result[f"PU_EXW__{slug}"] = sub["_PU_EXW"]
        result[f"Costo_Flete__{slug}"] = sub["_COSTO_FLETE"]
        result[f"PU_Total__{slug}"] = sub["_PU_TOTAL"]

        if subtotal_mode == "pu_total_x_cantidad":
            result[f"Subtotal__{slug}"] = sub["_PU_TOTAL"] * result["Cantidad_Total"]
        else:
            result[f"Subtotal__{slug}"] = None

        for dcol, fname in zip(dynamic_cols, dynamic_field_names):
            result[f"{fname}__{slug}"] = sub[dcol]

    # --- métricas globales por item ---
    pu_total_cols = [f"PU_Total__{provider_slug[p]}" for p in providers]

    result["PRECIO_MIN"] = result[pu_total_cols].min(axis=1, skipna=True)
    result["PRECIO_MAX"] = result[pu_total_cols].max(axis=1, skipna=True)
    result["DELTA_%"] = ((result["PRECIO_MAX"] - result["PRECIO_MIN"]) / result["PRECIO_MIN"] * 100).round(2)

    def _winner(row):
        best_p, best_v = None, None
        for p in providers:
            v = row.get(f"PU_Total__{provider_slug[p]}")
            if pd.notna(v) and (best_v is None or v < best_v):
                best_v, best_p = v, p
        return best_p

    result["PROVEEDOR_GANADOR"] = result.apply(_winner, axis=1)

    def _min_exw(row):
        winner = row["PROVEEDOR_GANADOR"]
        if pd.isna(winner) or winner not in provider_slug:
            return None
        return row.get(f"PU_EXW__{provider_slug[winner]}")

    result["PRECIO_MIN_EXW"] = result.apply(_min_exw, axis=1)

    result = result.reset_index(drop=True)
    return result, providers, dynamic_field_names


def process_file(file_path_or_buffer, sheet_name="Prop Adj Final"):
    df = load_prop_adj_final(file_path_or_buffer, sheet_name=sheet_name)
    consolidado, providers, dynamic_field_names = build_consolidado(df)
    return consolidado, providers, dynamic_field_names


# ===========================================================================
# DIRECCIÓN INVERSA: Consolidado (ancho) -> Prop Adj Final (largo)
# ===========================================================================

# Encabezados EXACTOS de la hoja 'Prop Adj Final' generada por Coupa.
# Se replican tal cual para que el archivo de salida tenga la misma forma
# que el original.
PROP_ADJ_FINAL_HEADERS = [
    "ID SAP",
    "País destino",
    "Descripción",
    "Cantidad Esperada",
    "Nombre del proveedor",
    "INCOTERM",
    "Factor Logistico",
    # <-- aquí se insertan los pares dinámicos por ronda:
    #     "Precio Unitario R{n}\n(USD)" + "Costo total  (USD) R{n}" -->
    "Mínimo costo\nofertado",
    "Desviación\ncon Mín [%]",
    "Calidad /\nHomologación",
    "Puntaje Técnico\n(/5.0)",
    # <-- aquí se insertan los campos descriptivos dinámicos detectados -->
    "Precio Histórico 2526",
    "Precio Histórico 2526 + factor",
    "Indice de inflación",
    "Baseline",
    "% Adj. E1",
    "Vol. E1",
    "Monto E1\n(USD)",
    "Ahorro_E1",
    "% Adj. E2",
    "Vol. E2",
    "Monto E2\n(USD)",
    "Ahorro_E2",
    "% Adj. E3\n← MANUAL",
    "Vol. E3",
    "Monto E3\n(USD)",
    "Ahorro_E3",
]

# Punto de inserción de los pares "Precio Unitario R{n}" / "Costo total R{n}".
_INSERT_RONDAS_AFTER = "Factor Logistico"

_INSERT_AFTER = "Puntaje Técnico\n(/5.0)"

# Campos dinámicos "conocidos" del Consolidado que sí tienen un slot fijo
# en Prop Adj Final (se mapean directo, no se duplican como columna extra).
_KNOWN_FIELD_SLOTS = {
    "homologacion": "Calidad /\nHomologación",
    "calidad_homologacion": "Calidad /\nHomologación",
}

# Campos del Consolidado que NUNCA deben tratarse como "dinámicos extra"
# porque ya se procesan aparte (precio) o son metadata de fila que no
# corresponde por proveedor.
_CORE_PRICE_FIELDS = {"PU_EXW", "Costo_Flete", "PU_Total", "Subtotal"}

# Columnas de métricas globales del Consolidado que no se trasladan fila a
# fila a Prop Adj Final (son resúmenes, no atributos ID+proveedor).
_GLOBAL_METRIC_COLS = {"PRECIO_MIN", "PRECIO_MAX", "DELTA_%", "PROVEEDOR_GANADOR", "PRECIO_MIN_EXW"}

_BASE_ITEM_COLS = ["ID", "Descripcion", "Pais", "Sociedad", "Partida_Principal",
                   "Subcategoria", "Codigo_SAP", "Cantidad_Total", "UM", "Zona_Lote"]


def _slot_key(field_name):
    return _norm(field_name).replace(" ", "_")


def _display_provider_name(token):
    """'Bermad_Perú_R2' -> 'Bermad Perú R2' ; 'DRIP-IRRI' se deja igual."""
    name = token.replace("_", " ").strip()
    name = re.sub(r"\s+", " ", name)
    return name


def load_consolidado(file_path_or_buffer, sheet_name=None):
    """Carga un Consolidado (formato ancho) desde Excel a DataFrame."""
    if sheet_name:
        return pd.read_excel(file_path_or_buffer, sheet_name=sheet_name)
    return pd.read_excel(file_path_or_buffer, sheet_name=0)


def parse_consolidado_columns(df):
    """
    Analiza las columnas del Consolidado y separa:
      - base_cols: columnas de identificación del item (sin '__')
      - per_provider: dict {proveedor: {campo: nombre_columna_real}}
      - extra_fields: lista de nombres de campo (sin duplicar) distintos a
        los core de precio (PU_EXW, Costo_Flete, PU_Total, Subtotal) y a
        los que ya tienen slot fijo (homologación, moneda, incoterm).
    """
    base_cols = []
    per_provider = {}
    extra_field_order = []

    for col in df.columns:
        if "__" not in col:
            if col not in _GLOBAL_METRIC_COLS:
                base_cols.append(col)
            continue
        field, provider_token = col.split("__", 1)
        provider = _display_provider_name(provider_token)
        per_provider.setdefault(provider, {})[field] = col

        if field not in _CORE_PRICE_FIELDS and _slot_key(field) not in _KNOWN_FIELD_SLOTS:
            if field not in extra_field_order:
                extra_field_order.append(field)

    return base_cols, per_provider, extra_field_order


def consolidado_to_prop_adj_final(df, rondas=(1,)):
    """
    Convierte un DataFrame en formato Consolidado (ancho) en un DataFrame
    con la estructura de 'Prop Adj Final' (largo: una fila por Item x Proveedor).

    'rondas': iterable con los números de ronda a representar como pares de
    columnas "Precio Unitario R{n}" / "Costo total (USD) R{n}". La última
    ronda de la lista es la ronda "vigente" (la que alimenta Mínimo,
    Desviación, Escenarios y Ahorro).

    Columnas que dependen de datos externos (Precio Histórico, Índice de
    inflación, % Adj. E3) quedan vacías para completarse a mano o vía cruce
    con otras hojas. Campos descriptivos por proveedor sin slot fijo
    (ej. Marca_Ofertada, Tiempo_Entrega, Garantia, Modalidad) se agregan como
    columnas extra después de 'Puntaje Técnico'.
    """
    base_cols, per_provider, extra_fields = parse_consolidado_columns(df)

    rondas = list(rondas) or [1]
    ronda_vigente = rondas[-1]

    # Construye la lista final de headers: base + pares por ronda + extras
    headers = []
    for h in PROP_ADJ_FINAL_HEADERS:
        headers.append(h)
        if h == _INSERT_RONDAS_AFTER:
            for n in rondas:
                headers.append(f"Precio Unitario R{n}\n(USD)")
                headers.append(f"Costo total  (USD) R{n}")
        if h == _INSERT_AFTER:
            headers.extend(extra_fields)

    id_col = "ID" if "ID" in df.columns else None
    desc_col = "Descripcion" if "Descripcion" in df.columns else None
    qty_col = "Cantidad_Total" if "Cantidad_Total" in df.columns else None
    pais_col = "Pais" if "Pais" in df.columns else None

    if id_col is None or desc_col is None or qty_col is None:
        raise ValueError(
            "El Consolidado debe tener al menos las columnas 'ID', 'Descripcion' "
            "y 'Cantidad_Total'."
        )

    providers = sorted(per_provider.keys())
    col_pu_vigente = f"Precio Unitario R{ronda_vigente}\n(USD)"

    rows = []
    for _, item in df.iterrows():
        item_id = item.get(id_col)
        if pd.isna(item_id):
            continue
        descripcion = item.get(desc_col)
        cantidad = item.get(qty_col)
        pais = item.get(pais_col) if pais_col else None

        for prov in providers:
            cols = per_provider[prov]
            pu_exw = item.get(cols["PU_EXW"]) if "PU_EXW" in cols else None
            costo_flete = item.get(cols["Costo_Flete"]) if "Costo_Flete" in cols else None
            pu_total = item.get(cols["PU_Total"]) if "PU_Total" in cols else None

            tiene_oferta = pd.notna(pu_total) and (pu_total or 0) > 0
            # Si el proveedor no tiene ningún dato para este item, se omite la fila
            if not tiene_oferta and all(pd.isna(item.get(cols.get(f))) for f in cols if f not in _CORE_PRICE_FIELDS):
                if pd.isna(pu_exw) and pd.isna(costo_flete) and pd.isna(pu_total):
                    continue

            row = {h: None for h in headers}
            row["ID SAP"] = item_id
            row["País destino"] = pais
            row["Descripción"] = descripcion
            row["Cantidad Esperada"] = cantidad
            row["Nombre del proveedor"] = prov

            # El "Precio Unitario" de la ronda vigente es el precio cotizado
            # (pre-factor logístico). El factor se aplica por fórmula en
            # 'Costo total (USD) R{n}'.
            precio_unitario = pu_total if pd.notna(pu_total) else pu_exw
            row[col_pu_vigente] = precio_unitario

            for slot_key, header in _KNOWN_FIELD_SLOTS.items():
                for field, real_col in cols.items():
                    if _slot_key(field) == slot_key:
                        row[header] = item.get(real_col)

            for field in extra_fields:
                if field in cols:
                    row[field] = item.get(cols[field])

            if not tiene_oferta and row.get("Calidad /\nHomologación") in (None, ""):
                row["Calidad /\nHomologación"] = "S/DATOS"

            rows.append(row)

    result = pd.DataFrame(rows, columns=headers)
    return result, providers, extra_fields


def process_consolidado_file(file_path_or_buffer, sheet_name=None, rondas=(1,)):
    df = load_consolidado(file_path_or_buffer, sheet_name=sheet_name)
    return consolidado_to_prop_adj_final(df, rondas=rondas)


# ===========================================================================
# Melteo COMPLETO del Consolidado (preserva TODOS los campos base: Pais,
# Sociedad, Partida_Principal, Subcategoria, Codigo_SAP, UM). Sirve de
# fuente para las hojas 'Detalle Ofertas', 'Detalle Evaluación Técnica' y
# 'Evaluación Económica'.
# ===========================================================================

DETALLE_OFERTAS_BASE_COLS = ["ID", "Descripcion", "Pais", "Sociedad", "Partida_Principal",
                             "Subcategoria", "Codigo_SAP", "Cantidad_Total", "UM", "Zona_Lote"]


def melt_consolidado_completo(df):
    """
    Convierte el Consolidado (ancho) en formato largo (una fila por Item +
    Proveedor), preservando TODOS los campos base disponibles (Pais,
    Sociedad, Partida_Principal, Subcategoria, Codigo_SAP, UM), a diferencia
    de 'consolidado_to_prop_adj_final' que sólo genera la estructura interna
    de Coupa.

    Devuelve (df_largo, providers) con columnas:
        ID, Descripcion, Pais, Sociedad, Partida_Principal, Subcategoria,
        Codigo_SAP, Cantidad_Total, UM, Zona_Lote, Proveedor,
        PU_EXW, Costo_Flete, PU_Total, Clave
    """
    base_cols, per_provider, _ = parse_consolidado_columns(df)

    id_col = "ID" if "ID" in df.columns else None
    if id_col is None:
        raise ValueError("El Consolidado debe tener una columna 'ID'.")

    providers = sorted(per_provider.keys())
    rows = []
    for _, item in df.iterrows():
        item_id = item.get(id_col)
        if pd.isna(item_id):
            continue

        base = {col: item.get(col) for col in DETALLE_OFERTAS_BASE_COLS if col in df.columns}

        for prov in providers:
            cols = per_provider[prov]
            pu_exw = item.get(cols["PU_EXW"]) if "PU_EXW" in cols else None
            costo_flete = item.get(cols["Costo_Flete"]) if "Costo_Flete" in cols else None
            pu_total = item.get(cols["PU_Total"]) if "PU_Total" in cols else None

            # Mismo criterio de inclusión que 'consolidado_to_prop_adj_final':
            # se incluye la fila si hay precio O si hay algún otro campo
            # (Marca, Garantía, etc.) informado para ese proveedor.
            tiene_oferta = pd.notna(pu_total) and (pu_total or 0) > 0
            tiene_otro_dato = any(
                pd.notna(item.get(real_col)) for f, real_col in cols.items()
                if f not in _CORE_PRICE_FIELDS
            )
            if not tiene_oferta and not tiene_otro_dato:
                if pd.isna(pu_exw) and pd.isna(costo_flete) and pd.isna(pu_total):
                    continue

            row = dict(base)
            row["Proveedor"] = prov
            row["PU_EXW"] = pu_exw
            row["Costo_Flete"] = costo_flete
            row["PU_Total"] = pu_total
            row["Clave"] = f"{item_id}|{prov}"
            rows.append(row)

    cols_final = DETALLE_OFERTAS_BASE_COLS + ["Proveedor", "PU_EXW", "Costo_Flete", "PU_Total", "Clave"]
    cols_final = [c for c in cols_final if c in (list(rows[0].keys()) if rows else cols_final)]
    result = pd.DataFrame(rows, columns=cols_final)
    return result, providers
