"""
Lectura del workbook de la ronda anterior.

Reemplaza los cinco campos de "archivos previos" que había antes (historial,
detalle técnico previo, evaluación económica previa, precio histórico y plan de
homologación) por UN solo archivo: el `PropAdjFinal_RondaN.xlsx` que la app
generó la vez pasada.

De ese archivo se extrae automáticamente:
  - historial de precios por ronda  (desde la hoja 'Prop Adj Final')
  - puntajes técnicos ya evaluados  (desde 'Detalle Evaluación Técnica')
  - precio histórico por ítem       (desde 'Evaluación Económica')
  - estados del plan de homologación (desde 'Plan Homologación', si existe)
"""

import re
from datetime import datetime

import pandas as pd

HOJA_PROP = "Prop Adj Final"
HOJA_TECNICA = "Detalle Evaluación Técnica"
HOJA_ECONOMICA = "Evaluación Económica"
HOJA_PLAN = "Plan Homologación"

# Las hojas de la app traen título en la fila 1-2 y encabezados en la fila 3.
FILA_ENCABEZADO = 2

COL_ID_PROP = "ID SAP"
COL_PROV_PROP = "Nombre del proveedor"
_RE_PRECIO_RONDA = re.compile(r"Precio Unitario R(\d+)", re.IGNORECASE)


def _leer_hoja(archivo, hoja, header=FILA_ENCABEZADO):
    try:
        if hasattr(archivo, "seek"):
            archivo.seek(0)
        df = pd.read_excel(archivo, sheet_name=hoja, header=header)
    except Exception:
        return None
    if df is None or df.empty:
        return None
    df = df.rename(columns=lambda c: str(c).strip())
    return df


def _historial_desde_prop(df_prop, archivo_origen):
    """
    Reconstruye el historial largo (ID | Proveedor | Ronda | Precio_DDP | ...)
    a partir de las columnas 'Precio Unitario R{n}' de la hoja Prop Adj Final.
    """
    if df_prop is None:
        return None, []

    cols_ronda = {}
    for c in df_prop.columns:
        m = _RE_PRECIO_RONDA.search(str(c).replace("\n", " "))
        if m:
            cols_ronda[int(m.group(1))] = c
    if not cols_ronda or COL_ID_PROP not in df_prop.columns:
        return None, []

    filas = []
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    for n in sorted(cols_ronda):
        col = cols_ronda[n]
        sub = df_prop[[COL_ID_PROP, COL_PROV_PROP, col]].copy()
        sub.columns = ["ID", "Proveedor", "Precio_DDP"]
        sub["Precio_DDP"] = pd.to_numeric(sub["Precio_DDP"], errors="coerce")
        sub = sub.dropna(subset=["ID", "Proveedor", "Precio_DDP"])
        sub = sub[sub["Precio_DDP"] > 0]
        if sub.empty:
            continue
        sub["ID"] = sub["ID"].astype(str).str.strip()
        sub["Proveedor"] = sub["Proveedor"].astype(str).str.strip()
        sub["Ronda"] = n
        sub["Archivo_Origen"] = archivo_origen
        sub["Fecha_Registro"] = ahora
        filas.append(sub[["ID", "Proveedor", "Ronda", "Precio_DDP",
                          "Archivo_Origen", "Fecha_Registro"]])

    if not filas:
        return None, []
    historial = pd.concat(filas, ignore_index=True)
    return historial, sorted(cols_ronda)


def leer_ronda_previa(archivo, nombre_archivo=""):
    """
    Devuelve (datos, resumen).

    datos = {
        'historial': DataFrame | None,
        'tecnica':   DataFrame | None,   # con la columna 'Clave'
        'economica': DataFrame | None,
        'precios_historicos': dict {ID: precio},
        'plan_estados': dict | None,
    }
    """
    datos = {"historial": None, "tecnica": None, "economica": None,
             "precios_historicos": {}, "plan_estados": None}
    resumen = {"hojas_leidas": [], "hojas_faltantes": [], "rondas": [],
               "registros_historial": 0, "puntajes": 0, "precios_historicos": 0,
               "estados_plan": 0}

    # --- Prop Adj Final -> historial de rondas ---
    df_prop = _leer_hoja(archivo, HOJA_PROP)
    if df_prop is not None:
        resumen["hojas_leidas"].append(HOJA_PROP)
        historial, rondas = _historial_desde_prop(df_prop, nombre_archivo)
        datos["historial"] = historial
        resumen["rondas"] = rondas
        resumen["registros_historial"] = 0 if historial is None else len(historial)
    else:
        resumen["hojas_faltantes"].append(HOJA_PROP)

    # --- Detalle Evaluación Técnica -> puntajes ya evaluados ---
    df_tec = _leer_hoja(archivo, HOJA_TECNICA)
    if df_tec is not None:
        resumen["hojas_leidas"].append(HOJA_TECNICA)
        df_tec = df_tec.rename(columns={"Clave (ID|Proveedor)": "Clave"})
        datos["tecnica"] = df_tec
        col_p = next((c for c in df_tec.columns if "puntaje" in c.lower()), None)
        if col_p:
            resumen["puntajes"] = int(pd.to_numeric(df_tec[col_p], errors="coerce").notna().sum())
    else:
        resumen["hojas_faltantes"].append(HOJA_TECNICA)

    # --- Evaluación Económica -> precio histórico por ítem ---
    df_eco = _leer_hoja(archivo, HOJA_ECONOMICA)
    if df_eco is not None:
        resumen["hojas_leidas"].append(HOJA_ECONOMICA)
        datos["economica"] = df_eco
        col_precio = next((c for c in df_eco.columns if "crudo" in c.lower()), None)
        if col_precio is None:
            col_precio = next(
                (c for c in df_eco.columns if "hist" in c.lower() and "precio" in c.lower()),
                None,
            )
        if col_precio and "ID" in df_eco.columns:
            precios = {}
            for _, r in df_eco.iterrows():
                v = pd.to_numeric(pd.Series([r.get(col_precio)]), errors="coerce").iloc[0]
                if pd.notna(v) and v > 0:
                    precios[str(r["ID"]).strip()] = float(v)
            datos["precios_historicos"] = precios
            resumen["precios_historicos"] = len(precios)
    else:
        resumen["hojas_faltantes"].append(HOJA_ECONOMICA)

    # --- Plan Homologación -> estados definidos a mano ---
    try:
        from .evaluacion import leer_plan_homologacion

        if hasattr(archivo, "seek"):
            archivo.seek(0)
        plan_estados, plan_resumen = leer_plan_homologacion(archivo)
        if plan_estados:
            datos["plan_estados"] = plan_estados
            resumen["estados_plan"] = plan_resumen.get("filas_leidas", 0)
            resumen["hojas_leidas"].append(HOJA_PLAN)
    except Exception:
        pass

    return datos, resumen


def puntajes_desde_tecnica(df_tecnica):
    """
    Convierte la hoja 'Detalle Evaluación Técnica' de una ronda previa en el
    dict {'ID|Proveedor': puntaje} que espera la etapa 3.
    """
    if df_tecnica is None or df_tecnica.empty:
        return {}
    col_clave = next((c for c in df_tecnica.columns if "clave" in c.lower()), None)
    col_punt = next((c for c in df_tecnica.columns if "puntaje" in c.lower()), None)
    if not col_clave or not col_punt:
        return {}
    out = {}
    for _, r in df_tecnica.iterrows():
        clave = str(r.get(col_clave) or "").strip()
        v = pd.to_numeric(pd.Series([r.get(col_punt)]), errors="coerce").iloc[0]
        if clave and "|" in clave and pd.notna(v) and v > 0:
            out[clave] = float(v)
    return out
