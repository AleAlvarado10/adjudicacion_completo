"""
Lectura del Precio Histórico desde un Excel externo (por ejemplo, la BD de importados).

No se le pregunta nada al usuario: recorre todas las hojas, detecta la fila de
encabezados, la columna del ID y la del precio, y se queda con la hoja que mejor
cruza con los ítems del proceso.

Está escrito para aguantar archivos reales: hojas con títulos en la primera fila,
encabezados duplicados o vacíos, columnas fantasma y hojas con el rango inflado.
"""

import re
import unicodedata

import pandas as pd

# Palabras que delatan cada columna, de más específica a más genérica.
PISTAS_ID = ["id sap", "idsap", "id", "codigo sap", "código sap", "material",
             "codigo", "código", "sku"]
PISTAS_PRECIO = ["precio historico", "precio histórico", "historico", "histórico",
                 "precio unitario", "precio", "costo unitario", "costo", "valor"]

MAX_FILAS_ENCABEZADO = 12


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def _columnas_unicas(valores):
    """Nombres de columna únicos y no vacíos (los repetidos quedan como 'Precio.2')."""
    salida, vistos = [], {}
    for i, v in enumerate(valores):
        nombre = "" if v is None else str(v).strip()
        if not nombre or nombre.lower() == "nan":
            nombre = f"col_{i}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}.{vistos[nombre]}"
        else:
            vistos[nombre] = 1
        salida.append(nombre)
    return salida


def _serie(df, col):
    """df[col] puede devolver un DataFrame si el nombre se repite: fuerza una Serie."""
    obj = df[col]
    if isinstance(obj, pd.DataFrame):
        obj = obj.iloc[:, 0]
    return obj


def _detectar_encabezado(df_raw):
    """La fila de encabezado es la que tiene más celdas de texto no vacías."""
    mejor, mejor_puntaje = 0, -1
    for i in range(min(MAX_FILAS_ENCABEZADO, len(df_raw))):
        fila = df_raw.iloc[i]
        puntaje = sum(1 for v in fila if isinstance(v, str) and v.strip())
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = i, puntaje
    return mejor


def _preparar(df_raw):
    """Limpia una hoja cruda y devuelve un DataFrame con encabezados utilizables."""
    df_raw = df_raw.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if df_raw.empty or df_raw.shape[1] < 2:
        return None
    df_raw = df_raw.reset_index(drop=True)
    fila_enc = _detectar_encabezado(df_raw)
    df = df_raw.iloc[fila_enc + 1:].reset_index(drop=True)
    if df.empty:
        return None
    df.columns = _columnas_unicas(df_raw.iloc[fila_enc].tolist())
    return df.dropna(axis=0, how="all")


def _candidatas(df, pistas, numerica=False):
    """Columnas que calzan con las pistas, en orden de preferencia."""
    encontradas = []
    normales = {c: _norm(c) for c in df.columns}
    for pista in pistas:
        for c, n in normales.items():
            if pista in n and c not in encontradas:
                if numerica and pd.to_numeric(_serie(df, c), errors="coerce").gt(0).sum() == 0:
                    continue
                encontradas.append(c)
    if numerica and not encontradas:
        # Último recurso: la columna numérica con más valores positivos.
        puntajes = []
        for c in df.columns:
            validos = int(pd.to_numeric(_serie(df, c), errors="coerce").gt(0).sum())
            if validos:
                puntajes.append((validos, c))
        if puntajes:
            encontradas = [max(puntajes)[1]]
    return encontradas


def _evaluar_hoja(df, ids_validos):
    """
    Prueba combinaciones de columna ID / columna precio y devuelve la mejor:
    (aciertos, col_id, col_precio, precios).
    """
    cols_id = _candidatas(df, PISTAS_ID)
    cols_precio = _candidatas(df, PISTAS_PRECIO, numerica=True)
    mejor = (0, None, None, {})

    for col_id in cols_id[:4]:
        # map(str) en vez de .str.strip(): la columna puede traer números o NaN
        ids = [str(v).strip() for v in _serie(df, col_id).tolist()]
        for col_precio in cols_precio[:4]:
            if col_precio == col_id:
                continue
            valores = pd.to_numeric(_serie(df, col_precio), errors="coerce")
            precios = {}
            for clave, valor in zip(ids, valores.tolist()):
                if not clave or clave.lower() in ("nan", "none"):
                    continue
                if pd.isna(valor) or valor <= 0:
                    continue
                precios[clave] = float(valor)
            if not precios:
                continue
            aciertos = (len([k for k in precios if k in ids_validos])
                        if ids_validos else len(precios))
            if aciertos > mejor[0]:
                mejor = (aciertos, col_id, col_precio, precios)
    return mejor


def leer_precio_historico_auto(archivo, ids_validos=None):
    """
    Devuelve (precios, resumen).
      precios: {ID: precio} — filtrado a los IDs del proceso si se pasan.
      resumen: {'hoja', 'col_id', 'col_precio', 'leidos', 'sin_match', 'hojas_revisadas'}
    """
    if hasattr(archivo, "seek"):
        archivo.seek(0)

    ids_validos = set(str(i).strip() for i in ids_validos) if ids_validos else set()
    hojas = pd.read_excel(archivo, sheet_name=None, header=None)

    mejor = (0, None, None, None, {})  # aciertos, hoja, col_id, col_precio, precios
    revisadas = []

    for nombre, df_raw in hojas.items():
        try:
            df = _preparar(df_raw)
        except Exception:
            continue
        if df is None:
            continue
        revisadas.append(nombre)
        try:
            aciertos, col_id, col_precio, precios = _evaluar_hoja(df, ids_validos)
        except Exception:
            continue
        if aciertos > mejor[0]:
            mejor = (aciertos, nombre, col_id, col_precio, precios)

    aciertos, nombre, col_id, col_precio, precios = mejor
    if not precios:
        raise ValueError(
            "No se encontró una columna con el ID y otra con el precio. Hojas "
            f"revisadas: {', '.join(revisadas) if revisadas else 'ninguna'}. "
            "Revisa que el archivo tenga el ID del material y su precio en la misma fila."
        )

    if ids_validos:
        del_proceso = {k: v for k, v in precios.items() if k in ids_validos}
        sin_match = len(precios) - len(del_proceso)
        precios = del_proceso
    else:
        sin_match = 0

    return precios, {
        "hoja": nombre, "col_id": col_id, "col_precio": col_precio,
        "leidos": len(precios), "sin_match": sin_match,
        "hojas_revisadas": len(revisadas),
    }
