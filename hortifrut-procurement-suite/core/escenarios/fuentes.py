"""
Lectura de FUENTES EXTERNAS con selección flexible de hoja y columna.

Reemplaza los XLOOKUP manuales que se hacían en Excel contra otros archivos,
por ejemplo:

    Precio Histórico = XLOOKUP(ID; 'Consolidado (2)'!A:A ; 'Consolidado (2)'!K:K)

Como la hoja y la columna cambian de un archivo a otro, aquí no se adivinan:
la app lista las hojas y las columnas disponibles y el usuario elige. Estas
funciones sólo hacen el cruce.

Dos fuentes soportadas:
  - Precio Histórico por ID           -> dict {ID: precio}
  - Puntaje Técnico por ID+Proveedor  -> dict {"ID|Proveedor": puntaje}
    (formato largo con columna de proveedor, o matriz con una columna por
     proveedor)
"""

import re
import unicodedata
import pandas as pd


def _norm(s):
    if s is None:
        return ""
    s = str(s).replace("\n", " ")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def listar_hojas(file_path_or_buffer):
    """Nombres de las hojas del archivo, en orden."""
    xls = pd.ExcelFile(file_path_or_buffer)
    return list(xls.sheet_names)


def _detectar_fila_encabezados(raw, max_scan=12):
    """
    Elige como fila de encabezados la primera fila de las primeras 'max_scan'
    que tenga al menos 2 celdas de texto no vacías y no numéricas.
    Devuelve el índice (0-based) o 0 si no encuentra nada mejor.
    """
    mejor, mejor_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        fila = raw.iloc[i].tolist()
        textos = sum(
            1 for v in fila
            if v is not None and not isinstance(v, (int, float))
            and str(v).strip() not in ("", "nan")
        )
        no_vacias = sum(1 for v in fila if v is not None and str(v).strip() not in ("", "nan"))
        if textos >= 2 and textos > mejor_score and no_vacias >= 2:
            mejor, mejor_score = i, textos
    return mejor


def leer_tabla(file_path_or_buffer, sheet_name, fila_encabezado=None):
    """
    Carga una hoja como DataFrame detectando (o recibiendo) la fila de
    encabezados. Devuelve (df, fila_encabezado_usada).
    """
    raw = pd.read_excel(file_path_or_buffer, sheet_name=sheet_name, header=None)
    if fila_encabezado is None:
        fila_encabezado = _detectar_fila_encabezados(raw)

    encabezados = []
    usados = {}
    for j, v in enumerate(raw.iloc[fila_encabezado].tolist()):
        nombre = str(v).strip() if v is not None and str(v).strip() not in ("", "nan") else f"Columna {j + 1}"
        if nombre in usados:
            usados[nombre] += 1
            nombre = f"{nombre} ({usados[nombre]})"
        else:
            usados[nombre] = 1
        encabezados.append(nombre)

    df = raw.iloc[fila_encabezado + 1:].copy()
    df.columns = encabezados
    df = df.dropna(how="all").reset_index(drop=True)
    return df, fila_encabezado


def sugerir_columna(columnas, *claves, numerica_en=None):
    """
    Devuelve la primera columna cuyo nombre normalizado contenga alguna de las
    'claves'. Si no encuentra y se pasa 'numerica_en' (un DataFrame), devuelve
    la primera columna mayormente numérica. Si nada aplica, None.
    """
    for c in columnas:
        cn = _norm(c)
        for k in claves:
            if k in cn:
                return c
    if numerica_en is not None:
        for c in columnas:
            serie = pd.to_numeric(numerica_en[c], errors="coerce")
            if serie.notna().sum() >= max(3, int(0.3 * len(serie))):
                return c
    return None


def leer_precio_historico(df, col_id, col_precio):
    """
    Devuelve (dict {ID: precio}, resumen).
    Ignora filas sin ID, sin precio, con precio 0 o no numérico.
    """
    precios, descartadas = {}, 0
    for _, r in df.iterrows():
        item = r.get(col_id)
        val = pd.to_numeric(r.get(col_precio), errors="coerce")
        if item is None or (isinstance(item, float) and pd.isna(item)) or str(item).strip() == "":
            continue
        if pd.isna(val) or val <= 0:
            descartadas += 1
            continue
        precios[str(item).strip()] = float(val)
    return precios, {"leidos": len(precios), "descartadas": descartadas}


def leer_puntajes_largo(df, col_id, col_proveedor, col_puntaje, escala_max=5.0):
    """
    Formato largo: una fila por ID + Proveedor.
    Devuelve (dict {"ID|Proveedor": puntaje}, resumen).
    """
    puntajes, descartadas, fuera_escala = {}, 0, 0
    for _, r in df.iterrows():
        item, prov = r.get(col_id), r.get(col_proveedor)
        val = pd.to_numeric(r.get(col_puntaje), errors="coerce")
        if any(v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""
               for v in (item, prov)):
            continue
        if pd.isna(val):
            descartadas += 1
            continue
        if val < 0 or val > escala_max:
            fuera_escala += 1
            continue
        puntajes[f"{str(item).strip()}|{str(prov).strip()}"] = float(val)
    return puntajes, {"leidos": len(puntajes), "descartadas": descartadas,
                      "fuera_escala": fuera_escala, "formato": "largo"}


def leer_puntajes_matriz(df, col_id, cols_proveedores, escala_max=5.0):
    """
    Formato matriz: una fila por ID, una columna por proveedor con el puntaje.
    Devuelve (dict {"ID|Proveedor": puntaje}, resumen).
    """
    puntajes, descartadas, fuera_escala = {}, 0, 0
    for _, r in df.iterrows():
        item = r.get(col_id)
        if item is None or (isinstance(item, float) and pd.isna(item)) or str(item).strip() == "":
            continue
        for pc in cols_proveedores:
            val = pd.to_numeric(r.get(pc), errors="coerce")
            if pd.isna(val):
                descartadas += 1
                continue
            if val < 0 or val > escala_max:
                fuera_escala += 1
                continue
            puntajes[f"{str(item).strip()}|{str(pc).strip()}"] = float(val)
    return puntajes, {"leidos": len(puntajes), "descartadas": descartadas,
                      "fuera_escala": fuera_escala, "formato": "matriz"}
