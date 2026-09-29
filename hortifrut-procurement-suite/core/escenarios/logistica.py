"""
Cruce con la tabla de Factor Logístico / Incoterm por Proveedor + País Destino.

Fuente esperada: Excel con una hoja que tenga columnas equivalentes a
    Proveedor | Pais de Origen | Ciudad de Origen | Incluido el costo logistico |
    Pais de Destino | Ciudad de Destino | Localidad / Comuna |
    Porcentaje o Factor logistico

Regla de negocio:
  - Si 'Incluido el costo logistico' = 'Incluido' -> INCOTERM = DDP.
    El costo cotizado ya integra flete: no se ajusta.
  - Si NO está marcado 'Incluido' -> INCOTERM = EXW, y se debe aplicar el
    'Porcentaje o Factor logistico' (ej. 0.32 = +32%) sobre el costo EXW
    cotizado para aproximar un costo DDP comparable:
        Costo_Ajustado = Costo_EXW * (1 + Factor)
  - Proveedores no encontrados en la tabla, o con múltiples rutas de
    factor distinto sin país de destino para desambiguar, quedan marcados
    para revisión manual (no se inventan factores).
"""

import re
import unicodedata
import pandas as pd


def _norm_key(s):
    """Normaliza para comparar: sin tildes, sin espacios/símbolos, mayúsculas."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    return s


def load_factor_logistico(file_path_or_buffer, sheet_name="Hoja1", header_row=1):
    """
    Carga la tabla de factor logístico y devuelve un DataFrame normalizado con
    columnas: Proveedor, Pais_Origen, Ciudad_Origen, Incluido, Pais_Destino,
    Ciudad_Destino, Localidad, Factor (+ columnas auxiliares _prov_norm,
    _destino_norm, _incluido_bool usadas para el matching).
    """
    df = pd.read_excel(file_path_or_buffer, sheet_name=sheet_name, header=header_row)
    df = df.rename(columns=lambda c: str(c).strip())

    col_map = {}
    for c in df.columns:
        cn = _norm_key(c)
        if cn == "PROVEEDOR":
            col_map[c] = "Proveedor"
        elif "PAISDEORIGEN" in cn:
            col_map[c] = "Pais_Origen"
        elif "CIUDADDEORIGEN" in cn:
            col_map[c] = "Ciudad_Origen"
        elif "INCLUIDO" in cn:
            col_map[c] = "Incluido"
        elif "PAISDEDESTINO" in cn:
            col_map[c] = "Pais_Destino"
        elif "CIUDADDEDESTINO" in cn or "CIUDADDEIMPORTACION" in cn:
            col_map[c] = "Ciudad_Destino"
        elif "LOCALIDAD" in cn or "COMUNA" in cn:
            col_map[c] = "Localidad"
        elif "FACTOR" in cn or "PORCENTAJE" in cn:
            col_map[c] = "Factor"

    df = df.rename(columns=col_map)
    keep = [c for c in ["Proveedor", "Pais_Origen", "Ciudad_Origen", "Incluido",
                         "Pais_Destino", "Ciudad_Destino", "Localidad", "Factor"]
            if c in df.columns]
    df = df[keep].dropna(subset=["Proveedor"]).reset_index(drop=True)

    df["_prov_norm"] = df["Proveedor"].apply(_norm_key)
    # El destino se compara contra País + Ciudad + Localidad concatenados, para
    # poder desambiguar por ciudad cuando el país solo no basta (ej. Perú-Trujillo
    # vs Perú-Olmos tienen factores distintos).
    partes_destino = [c for c in ["Pais_Destino", "Ciudad_Destino", "Localidad"] if c in df.columns]
    if partes_destino:
        df["_destino_norm"] = df[partes_destino].astype(str).agg("".join, axis=1).apply(_norm_key)
    else:
        df["_destino_norm"] = ""
    df["_incluido_bool"] = df["Incluido"].apply(
        lambda v: isinstance(v, str) and "INCLU" in _norm_key(v)
    )
    return df


def paises_disponibles(lookup_df):
    """Lista de países de destino presentes en la tabla (para el selector de default)."""
    if "Pais_Destino" not in lookup_df.columns:
        return []
    return sorted(lookup_df["Pais_Destino"].dropna().unique().tolist())


def match_factor(proveedor, pais_destino, lookup_df):
    """
    Busca la fila de 'lookup_df' que corresponde a (proveedor, pais_destino).
    Devuelve dict: {incoterm, factor, status, matched_row}
      status:
        "OK"                    -> resuelto sin ambigüedad
        "PROVEEDOR_NO_ENCONTRADO" -> el proveedor no está en la tabla
        "REQUIERE_DESTINO"      -> el proveedor tiene rutas con resultados
                                    distintos y no se pudo saber el destino
    """
    prov_norm = _norm_key(proveedor)
    dest_norm = _norm_key(pais_destino) if pais_destino else ""

    if not prov_norm:
        return {"incoterm": None, "factor": None, "status": "SIN_PROVEEDOR", "matched_row": None}

    candidatos = lookup_df[
        lookup_df["_prov_norm"].apply(lambda p: bool(p) and (p in prov_norm or prov_norm in p))
    ]

    if candidatos.empty:
        return {"incoterm": None, "factor": None, "status": "PROVEEDOR_NO_ENCONTRADO", "matched_row": None}

    if dest_norm:
        con_destino = candidatos[
            candidatos["_destino_norm"].apply(lambda d: bool(d) and (d in dest_norm or dest_norm in d))
        ]
        if not con_destino.empty:
            candidatos = con_destino

    # ¿Todas las filas candidatas dan el mismo resultado (mismo incoterm y factor)?
    factores_unicos = candidatos["Factor"].apply(lambda v: round(v, 6) if pd.notna(v) else None).nunique()
    incluido_unico = candidatos["_incluido_bool"].nunique()

    if factores_unicos > 1 or incluido_unico > 1:
        return {"incoterm": None, "factor": None, "status": "REQUIERE_DESTINO", "matched_row": None}

    row = candidatos.iloc[0]
    incluido = bool(row["_incluido_bool"])
    return {
        "incoterm": "DDP" if incluido else "EXW",
        "factor": None if incluido else row.get("Factor"),
        "status": "OK",
        "matched_row": row,
    }


def aplicar_factor_logistico(df, lookup_df, proveedor_col, destino_col=None,
                              costo_exw_col=None, costo_ddp_col=None,
                              destino_default=None):
    """
    Recorre 'df' (una fila por item+proveedor) y agrega:
      INCOTERM_Detectado, Factor_Logistico, Costo_Ajustado_Logistica, Estado_Logistico

    - costo_exw_col: costo base (EXW) sobre el que se aplica el factor cuando
      el proveedor NO incluye el costo logístico.
    - costo_ddp_col: costo ya calculado como landed/DDP; se usa tal cual
      cuando el proveedor SÍ incluye el costo logístico, o como respaldo si
      no hay match en la tabla.
    """
    incoterms, factores, ajustados, estados, destinos_usados = [], [], [], [], []

    for _, row in df.iterrows():
        proveedor = row.get(proveedor_col)
        destino = row.get(destino_col) if destino_col else None
        if (destino is None or (isinstance(destino, float) and pd.isna(destino)) or str(destino).strip() == "") \
                and destino_default:
            destino = destino_default

        m = match_factor(proveedor, destino, lookup_df)
        costo_exw = row.get(costo_exw_col) if costo_exw_col else None
        costo_ddp = row.get(costo_ddp_col) if costo_ddp_col else None

        incoterms.append(m["incoterm"])
        factores.append(m["factor"])
        estados.append(m["status"])
        destinos_usados.append(destino)

        base_fallback = costo_ddp if pd.notna(costo_ddp) else costo_exw

        if m["status"] != "OK":
            ajustados.append(base_fallback)
        elif m["incoterm"] == "DDP":
            ajustados.append(base_fallback)
        elif m["factor"] and pd.notna(costo_exw):
            ajustados.append(costo_exw * (1 + m["factor"]))
        else:
            ajustados.append(base_fallback)

    out = df.copy()
    out["Pais_Destino_Usado"] = destinos_usados
    out["INCOTERM_Detectado"] = incoterms
    out["Factor_Logistico"] = factores
    out["Costo_Ajustado_Logistica"] = ajustados
    out["Estado_Logistico"] = estados
    return out
