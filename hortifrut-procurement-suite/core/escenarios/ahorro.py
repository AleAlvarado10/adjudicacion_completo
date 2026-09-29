"""
Enriquece la hoja 'Prop Adj Final' con la estructura nueva:

  1. Factor logístico -> escribe en los slots 'INCOTERM' y 'Factor Logistico'.
     Regla: si el proveedor ya incluye el costo logístico en su precio ->
     INCOTERM = DDP y Factor = 0 (no se ajusta). Si no lo incluye ->
     INCOTERM = EXW y Factor = el de la ruta (Proveedor + País destino).
     El ajuste en sí lo hace la fórmula de Excel:
         Costo total (USD) R{n} = Precio Unitario R{n} × (1 + Factor Logistico)

  2. Historial de rondas -> llena 'Precio Unitario R{n} (USD)' de las rondas
     anteriores y registra el precio unitario de la ronda actual.

  3. Índice de inflación (parámetro) y Baseline:
         Precio Histórico + factor = Precio Histórico × (1 + Factor Logistico)
         Baseline = (Precio Histórico + factor) × (1 + Índice de inflación)
                    y si no hay Precio Histórico -> Costo total de la Ronda 1
         Ahorro_E{n} = (Baseline − Costo total ronda vigente) × Vol. E{n}

Las columnas de auditoría logística (Pais_Destino_Usado, Factor_Logistico,
Estado_Logistico) se calculan pero NO van a la hoja final: se usan sólo para
los avisos de la app. Se eliminan con `quitar_columnas_auditoria`.
"""

import pandas as pd

from .logistica import aplicar_factor_logistico
from .rondas import actualizar_historial, historial_wide

COL_ID = "ID SAP"
COL_PROVEEDOR = "Nombre del proveedor"
COL_PAIS = "País destino"
COL_CANTIDAD = "Cantidad Esperada"
COL_INCOTERM = "INCOTERM"
COL_FACTOR = "Factor Logistico"
COL_HISTORICO = "Precio Histórico 2526"
COL_HISTORICO_FACTOR = "Precio Histórico 2526 + factor"
COL_INFLACION = "Indice de inflación"
COL_BASELINE = "Baseline"

AUDIT_COLS = ["Pais_Destino_Usado", "Factor_Logistico", "Estado_Logistico"]

# Columnas que solo tienen sentido si se subió la tabla de Factor Logístico.
# Si no se subió, se eliminan de la hoja final (ver `quitar_columnas_logistica`).
LOGISTICA_COLS = [COL_INCOTERM, COL_FACTOR, COL_HISTORICO_FACTOR]


def col_precio_unitario(n):
    return f"Precio Unitario R{n}\n(USD)"


def col_costo_total(n):
    return f"Costo total  (USD) R{n}"


def enriquecer(df, lookup_logistico=None, destino_default=None,
               historial_df=None, ronda_actual=None, archivo_origen="",
               indice_inflacion=0.0, rondas=None):
    """
    Devuelve (df_enriquecido, historial_actualizado, resumen_dict).

    'rondas': lista de rondas representadas en el DataFrame (1..N). La última
    es la ronda vigente.
    """
    out = df.copy()
    resumen = {}

    rondas = list(rondas) if rondas else [ronda_actual or 1]
    ronda_vigente = rondas[-1]

    # --- 1. Factor logístico -> INCOTERM + Factor Logistico ---
    destino_col = COL_PAIS if COL_PAIS in out.columns else None
    pu_vigente = col_precio_unitario(ronda_vigente)

    if lookup_logistico is not None:
        out = aplicar_factor_logistico(
            out, lookup_logistico,
            proveedor_col=COL_PROVEEDOR,
            destino_col=destino_col,
            costo_exw_col=pu_vigente if pu_vigente in out.columns else None,
            costo_ddp_col=pu_vigente if pu_vigente in out.columns else None,
            destino_default=destino_default,
        )
        resumen["estados_logistico"] = out["Estado_Logistico"].value_counts(dropna=False).to_dict()
        out[COL_INCOTERM] = out["INCOTERM_Detectado"]
        # Si el proveedor ya incluye logística (DDP) el factor es 0; si no hay
        # ruta en la tabla, también 0 (no se inventa un ajuste).
        out[COL_FACTOR] = out["Factor_Logistico"].fillna(0)
        out = out.drop(columns=["INCOTERM_Detectado", "Costo_Ajustado_Logistica"], errors="ignore")
    else:
        out[COL_INCOTERM] = None
        out[COL_FACTOR] = 0
        out["Pais_Destino_Usado"] = None
        out["Factor_Logistico"] = 0
        out["Estado_Logistico"] = "SIN_TABLA_LOGISTICA"

    # --- 2. Historial de rondas: precio unitario por ronda ---
    if ronda_actual is not None and pu_vigente in out.columns:
        historial_df = actualizar_historial(
            historial_df, out, ronda=ronda_actual,
            id_col=COL_ID, proveedor_col=COL_PROVEEDOR,
            precio_col=pu_vigente, archivo_origen=archivo_origen,
        )

    wide = historial_wide(historial_df)
    if not wide.empty:
        wide = wide.copy()
        wide["_ID"] = wide["ID"].astype(str)
        wide["_PROV"] = wide["Proveedor"].astype(str)
        idx = wide.set_index(["_ID", "_PROV"])
        claves = list(zip(out[COL_ID].astype(str), out[COL_PROVEEDOR].astype(str)))
        for n in rondas:
            if n == ronda_vigente:
                continue  # la ronda vigente ya viene del Consolidado subido
            origen = f"Precio_Ronda_{n}"
            destino = col_precio_unitario(n)
            if origen in idx.columns and destino in out.columns:
                out[destino] = [
                    idx[origen].get(k) if k in idx.index else None for k in claves
                ]

    # --- 3. Índice de inflación: valor único o mapa por país (tabla CPI) ---
    if COL_INFLACION in out.columns:
        if isinstance(indice_inflacion, dict):
            from .inflacion import factor_para, paises_no_resueltos

            paises = out[COL_PAIS] if COL_PAIS in out.columns else pd.Series([None] * len(out))
            valores = [factor_para(indice_inflacion, p) for p in paises]
            resumen["inflacion_resueltas"] = sum(1 for v in valores if v is not None)
            resumen["inflacion_sin_pais"] = paises_no_resueltos(indice_inflacion, paises.tolist())
            out[COL_INFLACION] = [0 if v is None else v for v in valores]
        else:
            out[COL_INFLACION] = indice_inflacion

    resumen["filas"] = len(out)
    resumen["rondas"] = rondas
    resumen["ronda_vigente"] = ronda_vigente
    return out, historial_df, resumen


def quitar_columnas_logistica(df):
    """
    Devuelve el DataFrame sin las columnas del factor logístico
    (INCOTERM, Factor Logistico y 'Precio Histórico 2526 + factor').
    Se usa cuando el usuario NO subió la tabla de Factor Logístico: en ese caso
    esas columnas solo agregaban ruido con valores vacíos o en cero.
    """
    return df.drop(columns=[c for c in LOGISTICA_COLS if c in df.columns], errors="ignore")


def quitar_columnas_auditoria(df):
    """Devuelve el DataFrame sin las columnas de auditoría logística
    (se usan sólo para los avisos de la app, no van a la hoja final)."""
    return df.drop(columns=[c for c in AUDIT_COLS if c in df.columns], errors="ignore")
