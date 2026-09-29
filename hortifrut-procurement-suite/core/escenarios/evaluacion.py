"""
Construye los esqueletos de dos hojas de entrada MANUAL:

  - 'Detalle Evaluación Técnica': una fila por Item + Proveedor, con
    'Puntaje Técnico (/5.0)' en blanco para que el equipo técnico lo
    complete. Semáforo y Estado Homologación se calculan por FÓRMULA a
    partir de umbrales editables (por defecto Alto ≥3.75, Medio ≥2.5).

  - 'Evaluación Económica': una fila por Item (no por proveedor), con
    'Precio Histórico' en blanco/0 para completar a mano. Alimenta el
    'Precio Histórico' de 'Prop Adj Final' vía cruce por ID.

Ambas soportan "carry-forward": si se sube la versión de una ronda
anterior, los puntajes / precios ya ingresados se recuperan por clave
(Clave = ID|Proveedor, o ID) en vez de partir en blanco de nuevo.
"""

import pandas as pd

UMBRAL_ALTO_DEFAULT = 3.75
UMBRAL_MEDIO_DEFAULT = 2.5


def construir_detalle_evaluacion_tecnica(df_melted, historial_tecnico=None, plan_estados=None,
                                          puntajes_externos=None):
    """
    df_melted: salida de transformer.melt_consolidado_completo (con columna 'Clave').
    historial_tecnico: DataFrame opcional (de una ronda anterior) con columnas
        ['Clave', 'Puntaje Técnico (/5.0)'] para recuperar puntajes ya evaluados.
    plan_estados: dict opcional {Clave: estado} proveniente de un Plan de
        Homologación importado. Tiene prioridad sobre el puntaje técnico.
    """
    base = (df_melted[["Clave", "ID", "Proveedor"]] if "ID" in df_melted.columns
            else df_melted[["Clave", "Proveedor"]])
    base = base.drop_duplicates(subset=["Clave"]).reset_index(drop=True)

    puntajes = {}
    if historial_tecnico is not None and not historial_tecnico.empty:
        col_puntaje = next((c for c in historial_tecnico.columns if "puntaje" in c.lower()), None)
        if col_puntaje and "Clave" in historial_tecnico.columns:
            for _, r in historial_tecnico.iterrows():
                v = r.get(col_puntaje)
                if pd.notna(v):
                    puntajes[r["Clave"]] = v

    # Un archivo externo de puntajes tiene prioridad sobre el historial de la
    # ronda anterior (es el dato más fresco que entrega el equipo técnico).
    if puntajes_externos:
        puntajes.update({k: v for k, v in puntajes_externos.items() if pd.notna(v)})

    base["Puntaje Técnico (/5.0)"] = base["Clave"].map(puntajes)
    base["Estado desde Plan"] = base["Clave"].map(plan_estados or {})

    cols = ["Clave", "ID", "Proveedor", "Puntaje Técnico (/5.0)", "Estado desde Plan"]
    return base[[c for c in cols if c in base.columns]]


# ---------------------------------------------------------------------------
# Lector flexible del Plan de Homologación
# ---------------------------------------------------------------------------

ESTADOS_VALIDOS = {
    "HOMOLOGADO": "HOMOLOGADO",
    "APROBADO": "HOMOLOGADO",
    "SI": "HOMOLOGADO",
    "SÍ": "HOMOLOGADO",
    "OK": "HOMOLOGADO",
    "POR REVISAR": "POR REVISAR",
    "EN REVISION": "POR REVISAR",
    "EN REVISIÓN": "POR REVISAR",
    "PENDIENTE": "POR REVISAR",
    "NO HOMOLOGADO": "NO HOMOLOGADO",
    "RECHAZADO": "NO HOMOLOGADO",
    "NO": "NO HOMOLOGADO",
    "OBSOLETO": "NO HOMOLOGADO",
}


def _normalizar_estado(valor):
    """Devuelve HOMOLOGADO / POR REVISAR / NO HOMOLOGADO, o None si no aplica."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    txt = str(valor).strip().upper()
    if not txt or txt in ("S/DATOS", "SIN DATOS", "-", "NAN"):
        return None
    return ESTADOS_VALIDOS.get(txt)


def leer_plan_homologacion(file_path_or_buffer, sheet_name=None):
    """
    Lee un Plan de Homologación y devuelve (dict {Clave: estado}, resumen_dict).

    Acepta dos formatos, detectados automáticamente:

      A) MATRIZ  — una fila por ID, una columna por proveedor:
             ID | Descripción | Proveedor A | Proveedor B | ...
         (es el formato que genera esta misma app en la hoja 'Plan Homologación')

      B) LARGO   — una fila por combinación:
             ID | Proveedor | Estado        (los nombres de columna son flexibles)

    Busca la fila de encabezados en las primeras filas, así que tolera títulos
    y notas arriba de la tabla.
    """
    hojas = pd.read_excel(file_path_or_buffer, sheet_name=sheet_name, header=None)
    if isinstance(hojas, dict):
        # Prioriza una hoja que se llame como el plan; si no, la primera.
        clave = next((k for k in hojas if "homolog" in str(k).lower()), list(hojas)[0])
        raw = hojas[clave]
    else:
        raw = hojas

    # --- Encontrar la fila de encabezados: la primera que contenga "ID" ---
    header_idx = None
    for i in range(min(10, len(raw))):
        fila = [str(v).strip().upper() for v in raw.iloc[i].tolist()]
        if any(v == "ID" or v.endswith("ID SAP") or v == "ID SAP" for v in fila):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(
            "No se encontró una fila de encabezados con una columna 'ID' en el "
            "Plan de Homologación."
        )

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = [str(v).strip() if v is not None else "" for v in raw.iloc[header_idx].tolist()]
    df = df.loc[:, [c for c in df.columns if c and c.lower() != "nan"]]

    def _find(*claves):
        for c in df.columns:
            cn = str(c).strip().lower()
            if any(k in cn for k in claves):
                return c
        return None

    col_id = _find("id sap", "id")
    col_prov = _find("proveedor")
    col_estado = _find("estado", "homologaci")

    estados = {}
    resumen = {"formato": None, "filas_leidas": 0, "descartadas": 0}

    # --- Formato LARGO: hay columna de proveedor Y de estado ---
    if col_id and col_prov and col_estado:
        resumen["formato"] = "largo"
        for _, r in df.iterrows():
            item = r.get(col_id)
            prov = r.get(col_prov)
            est = _normalizar_estado(r.get(col_estado))
            if pd.isna(item) or pd.isna(prov):
                continue
            if est is None:
                resumen["descartadas"] += 1
                continue
            estados[f"{str(item).strip()}|{str(prov).strip()}"] = est
            resumen["filas_leidas"] += 1
        return estados, resumen

    # --- Formato MATRIZ: cada columna que no sea ID/Descripción es un proveedor ---
    if col_id:
        resumen["formato"] = "matriz"
        col_desc = _find("descripci")
        prov_cols = [c for c in df.columns if c not in (col_id, col_desc)]
        for _, r in df.iterrows():
            item = r.get(col_id)
            if pd.isna(item):
                continue
            for pc in prov_cols:
                est = _normalizar_estado(r.get(pc))
                if est is None:
                    resumen["descartadas"] += 1
                    continue
                estados[f"{str(item).strip()}|{str(pc).strip()}"] = est
                resumen["filas_leidas"] += 1
        resumen["proveedores"] = [str(c).strip() for c in prov_cols]
        return estados, resumen

    raise ValueError(
        "El Plan de Homologación debe tener una columna 'ID' y, o bien columnas "
        "'Proveedor' + 'Estado', o bien una columna por proveedor."
    )


def construir_evaluacion_economica(df_melted, historial_economico=None, precios_externos=None):
    """
    historial_economico: DataFrame opcional (de una ronda anterior) con columnas
        ['ID', 'Precio Histórico (crudo)'] para recuperar precios ya ingresados.
    """
    base = df_melted[["ID", "Descripcion"]].drop_duplicates(subset=["ID"]).reset_index(drop=True)

    precios = {}
    if historial_economico is not None and not historial_economico.empty:
        col_precio = next((c for c in historial_economico.columns if "crudo" in c.lower()), None)
        if col_precio and "ID" in historial_economico.columns:
            for _, r in historial_economico.iterrows():
                v = r.get(col_precio)
                if pd.notna(v) and v != 0:
                    precios[r["ID"]] = v

    # Un archivo externo de precio histórico tiene prioridad sobre lo que
    # venía de la ronda anterior.
    if precios_externos:
        precios.update({k: v for k, v in precios_externos.items() if pd.notna(v)})

    base["Precio Histórico (crudo)"] = base["ID"].astype(str).str.strip().map(precios).fillna(0)
    return base


def construir_plan_homologacion(df_melted, providers):
    """Esqueleto ID x Descripción (filas) x Proveedor (columnas), vacío —
    las celdas se llenan con fórmulas INDEX/MATCH al escribir la hoja."""
    base = df_melted[["ID", "Descripcion"]].drop_duplicates(subset=["ID"]).reset_index(drop=True)
    for p in providers:
        base[p] = None
    return base
