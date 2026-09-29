"""
Índice de inflación (CPI) por país.

Fuente esperada: un Excel con dos columnas — país y tasa. La tasa se acepta
como porcentaje en texto ("2,1%" o "2.1%"), como número entero de porcentaje
(2.1) o como fracción (0.021). Todo se normaliza a fracción: 0.021.

Se usa para llenar la columna 'Indice de inflación' de la hoja Prop Adj Final
según el 'País destino' de cada ítem, en vez de un único valor manual.
"""

import re
import unicodedata

import pandas as pd

# Alias frecuentes: lo que dice la tabla CPI vs. lo que traen las cotizaciones.
ALIAS = {
    "USA": "ESTADOSUNIDOS",
    "US": "ESTADOSUNIDOS",
    "EEUU": "ESTADOSUNIDOS",
    "EE UU": "ESTADOSUNIDOS",
    "UNITEDSTATES": "ESTADOSUNIDOS",
    "UK": "UNITEDKINGDOM",
    "INGLATERRA": "UNITEDKINGDOM",
    "REINOUNIDO": "UNITEDKINGDOM",
    "BRASIL": "BRAZIL",
    "PAISESBAJOS": "HOLANDA",
    "NETHERLANDS": "HOLANDA",
    "SPAIN": "ESPANA",
    "MOROCCO": "MARRUECOS",
    "IRELAND": "IRLANDA",
}


def _norm(s):
    """Sin tildes, sin espacios ni símbolos, mayúsculas."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def _a_fraccion(valor):
    """'2,1%' -> 0.021 · 2.1 -> 0.021 · 0.021 -> 0.021 · None si no se puede."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, str):
        txt = valor.strip().replace("%", "").replace(",", ".")
        try:
            num = float(txt)
        except ValueError:
            return None
        # Venía con '%' o es un número grande: está en puntos porcentuales.
        return num / 100 if ("%" in valor or num > 1) else num
    try:
        num = float(valor)
    except (TypeError, ValueError):
        return None
    return num / 100 if num > 1 else num


def load_inflacion(file_path_or_buffer, sheet_name=0):
    """
    Devuelve (mapa, resumen).
      mapa: {PAIS_NORMALIZADO: fraccion}
      resumen: {'paises': n, 'descartadas': n, 'detalle': [(pais, fraccion), ...]}
    Detecta la fila de encabezado sola: si la primera fila ya es un país con
    tasa, la incluye igual (la tabla CPI no siempre trae títulos).
    """
    if hasattr(file_path_or_buffer, "seek"):
        file_path_or_buffer.seek(0)
    df = pd.read_excel(file_path_or_buffer, sheet_name=sheet_name, header=None)
    if df.shape[1] < 2:
        raise ValueError(
            "La tabla de inflación necesita al menos dos columnas: país y tasa."
        )

    mapa, detalle, descartadas = {}, [], 0
    for _, row in df.iterrows():
        pais = _norm(row.iloc[0])
        frac = _a_fraccion(row.iloc[1])
        if not pais or frac is None:
            descartadas += 1
            continue
        # Descarta la fila de encabezado (ej. 'CPI' | 'T26/27')
        if pais in {"CPI", "PAIS", "COUNTRY", "PAISES"}:
            descartadas += 1
            continue
        mapa[pais] = frac
        detalle.append((str(row.iloc[0]).strip(), frac))

    if not mapa:
        raise ValueError(
            "No se pudo leer ninguna tasa. Revisa que la primera columna sea el "
            "país y la segunda el porcentaje (ej. PERU | 2,2%)."
        )

    return mapa, {"paises": len(mapa), "descartadas": descartadas, "detalle": detalle}


def factor_para(mapa, pais):
    """Busca el índice de un país, tolerando alias y variantes de escritura."""
    clave = _norm(pais)
    if not clave:
        return None
    clave = ALIAS.get(clave, clave)
    if clave in mapa:
        return mapa[clave]
    # Coincidencia parcial: 'PERU' vs 'PERUANO', 'CHILE' vs 'CHILECENTRAL'
    for k, v in mapa.items():
        if k and (k in clave or clave in k):
            return v
    return None


def paises_no_resueltos(mapa, paises):
    """Lista de países presentes en los datos que no están en la tabla CPI."""
    faltan = []
    for p in paises:
        if p is None or (isinstance(p, float) and pd.isna(p)):
            continue
        if factor_para(mapa, p) is None and str(p).strip():
            faltan.append(str(p).strip())
    return sorted(set(faltan))
