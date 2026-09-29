"""
Historial de precios por Ronda de cotización.

Cada vez que se procesa un archivo (Ronda 1, 2, 3...), se registra el precio
ajustado (DDP) por ID + Proveedor. Esto permite:
  - Comparar Precio_Ronda_1 vs Precio_Ronda_2 vs Precio_Ronda_3...
  - Usar el precio de la Ronda 1 como respaldo de 'Precio Histórico' cuando
    un ítem no tiene precio histórico real.

Persistencia: se guarda en un Excel local (historial_rondas.xlsx) en la
carpeta donde corre la app. Además, se puede subir/descargar manualmente
el historial para llevarlo entre computadores o sesiones.
"""

import os
import re
from datetime import datetime
import pandas as pd

HISTORIAL_COLUMNS = ["ID", "Proveedor", "Ronda", "Precio_DDP", "Archivo_Origen", "Fecha_Registro"]
DEFAULT_HISTORIAL_PATH = "historial_rondas.xlsx"


def detectar_numero_ronda(filename):
    """Busca un patrón 'Ronda 2', 'Ronda_2', 'R2' en el nombre del archivo."""
    if not filename:
        return None
    m = re.search(r"ronda[\s_-]*([0-9]+)", filename, flags=re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r"[^A-Za-z](R)([0-9]+)[^0-9]", " " + filename + " ")
    if m:
        return int(m.group(2))
    return None


def cargar_historial(uploaded_file=None):
    """
    Carga el historial existente.
    Prioridad: archivo subido manualmente > archivo local en disco > vacío.
    """
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Historial")
            for col in HISTORIAL_COLUMNS:
                if col not in df.columns:
                    df[col] = None
            return df[HISTORIAL_COLUMNS]
        except Exception:
            pass

    if os.path.exists(DEFAULT_HISTORIAL_PATH):
        try:
            df = pd.read_excel(DEFAULT_HISTORIAL_PATH, sheet_name="Historial")
            for col in HISTORIAL_COLUMNS:
                if col not in df.columns:
                    df[col] = None
            return df[HISTORIAL_COLUMNS]
        except Exception:
            pass

    return pd.DataFrame(columns=HISTORIAL_COLUMNS)


def actualizar_historial(historial_df, nuevos_df, ronda, id_col, proveedor_col, precio_col, archivo_origen=""):
    """
    Inserta/actualiza (upsert) los precios de 'nuevos_df' para la ronda indicada.
    Si ya existía un registro (mismo ID+Proveedor+Ronda), se reemplaza (no se duplica).
    """
    nuevos = nuevos_df[[id_col, proveedor_col, precio_col]].dropna(subset=[id_col, proveedor_col]).copy()
    nuevos = nuevos.rename(columns={id_col: "ID", proveedor_col: "Proveedor", precio_col: "Precio_DDP"})
    nuevos["ID"] = nuevos["ID"].astype(str)
    nuevos["Proveedor"] = nuevos["Proveedor"].astype(str)
    nuevos["Ronda"] = ronda
    nuevos["Archivo_Origen"] = archivo_origen
    nuevos["Fecha_Registro"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    nuevos = nuevos[HISTORIAL_COLUMNS]

    if historial_df is None or historial_df.empty:
        return nuevos.reset_index(drop=True)

    historial_df = historial_df.copy()
    historial_df["ID"] = historial_df["ID"].astype(str)
    historial_df["Proveedor"] = historial_df["Proveedor"].astype(str)

    llave_nueva = set(zip(nuevos["ID"], nuevos["Proveedor"], nuevos["Ronda"]))
    mask_mantener = ~historial_df.apply(
        lambda r: (r["ID"], r["Proveedor"], r["Ronda"]) in llave_nueva, axis=1
    )
    combinado = pd.concat([historial_df[mask_mantener], nuevos], ignore_index=True)
    return combinado


def historial_wide(historial_df):
    """Pivotea a formato ancho: ID, Proveedor, Precio_Ronda_1, Precio_Ronda_2, ..."""
    if historial_df is None or historial_df.empty:
        return pd.DataFrame(columns=["ID", "Proveedor"])
    pivot = historial_df.pivot_table(
        index=["ID", "Proveedor"], columns="Ronda", values="Precio_DDP", aggfunc="last"
    )
    pivot.columns = [f"Precio_Ronda_{int(c)}" for c in pivot.columns]
    pivot = pivot.reset_index()
    return pivot


def guardar_historial_local(historial_df, path=DEFAULT_HISTORIAL_PATH):
    """Intenta guardar en disco local (best-effort); devuelve True/False según éxito."""
    try:
        historial_df.to_excel(path, sheet_name="Historial", index=False)
        return True
    except Exception:
        return False


def precio_ronda_1(historial_df, item_id, proveedor):
    """Devuelve el precio registrado en Ronda 1 para (item_id, proveedor), o None."""
    if historial_df is None or historial_df.empty:
        return None
    sub = historial_df[
        (historial_df["ID"].astype(str) == str(item_id))
        & (historial_df["Proveedor"].astype(str) == str(proveedor))
        & (historial_df["Ronda"] == 1)
    ]
    if sub.empty:
        return None
    return sub.iloc[-1]["Precio_DDP"]
