"""
Diagramas del flujo de trabajo, uno por módulo.

SVG plano (sin dependencias) con la paleta Hortifrut. Se renderiza con
st.markdown(..., unsafe_allow_html=True) desde la página de inicio.
"""

from . import constantes as C

# Paleta de los diagramas
_ENTRADA_BG, _ENTRADA_BR, _ENTRADA_TX = "#F1EFE8", "#B4B2A9", "#2C2C2A"
_PROCESO_BG, _PROCESO_BR, _PROCESO_TX = "#EEEDFE", C.AZUL_CLARO, C.AZUL
_SALIDA_BG, _SALIDA_BR, _SALIDA_TX = "#E1F5EE", "#1D9E75", "#0F6E56"
_SESION_BG, _SESION_BR, _SESION_TX = "#FAECE7", "#D85A30", "#993C1D"
_ESC_BG, _ESC_BR, _ESC_TX = "#FAEEDA", "#EF9F27", "#854F0B"
_LINEA = "#888780"
_FUENTE = '"Century Gothic", Poppins, Futura, sans-serif'

_DEFS = (
    '<defs><marker id="hf-arrow" viewBox="0 0 10 10" refX="8" refY="5" '
    'markerWidth="6" markerHeight="6" orient="auto-start-reverse">'
    f'<path d="M2 1L8 5L2 9" fill="none" stroke="{_LINEA}" stroke-width="1.5" '
    'stroke-linecap="round" stroke-linejoin="round"/></marker></defs>'
)


def _caja(x, y, w, h, titulo, sub, bg, br, tx, size=13):
    cx = x + w / 2
    out = (
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="8" fill="{bg}" '
        f'stroke="{br}" stroke-width="1"/>'
    )
    if sub:
        out += (
            f'<text x="{cx}" y="{y + h / 2 - 9}" text-anchor="middle" '
            f'dominant-baseline="central" font-family={_FUENTE!r} font-size="{size}" '
            f'font-weight="600" fill="{tx}">{titulo}</text>'
            f'<text x="{cx}" y="{y + h / 2 + 10}" text-anchor="middle" '
            f'dominant-baseline="central" font-family={_FUENTE!r} font-size="11" '
            f'fill="{tx}" opacity="0.75">{sub}</text>'
        )
    else:
        out += (
            f'<text x="{cx}" y="{y + h / 2}" text-anchor="middle" '
            f'dominant-baseline="central" font-family={_FUENTE!r} font-size="{size}" '
            f'font-weight="600" fill="{tx}">{titulo}</text>'
        )
    return out


def _flecha(x1, y1, x2, y2):
    return (
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{_LINEA}" '
        f'stroke-width="1.2" marker-end="url(#hf-arrow)"/>'
    )


def _contenedor(x, y, w, h, etiqueta="Proceso"):
    return (
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="12" fill="none" '
        f'stroke="{_LINEA}" stroke-width="0.8" stroke-dasharray="4 4" opacity="0.7"/>'
        f'<text x="{x + 16}" y="{y + 18}" font-family={_FUENTE!r} font-size="11" '
        f'fill="{_LINEA}">{etiqueta}</text>'
    )


def _svg(alto, cuerpo, titulo):
    return (
        f'<svg width="100%" viewBox="0 0 680 {alto}" role="img" '
        f'xmlns="http://www.w3.org/2000/svg" style="max-width:680px">'
        f"<title>{titulo}</title>{_DEFS}{cuerpo}</svg>"
    )


# Columnas reutilizadas por los tres diagramas
_X = (56, 252, 448)
_W = 176


def consolidador() -> str:
    c = _caja(230, 10, 220, 52, "Excel de proveedores", "uno por proveedor",
              _ENTRADA_BG, _ENTRADA_BR, _ENTRADA_TX)
    c += _flecha(340, 62, 340, 92)
    c += _contenedor(40, 96, 600, 106)
    pasos = [("Normaliza", "19 campos por ítem"),
             ("Tabla comparativa", "ítem × proveedor"),
             ("Benchmarking", "mín, máx, delta %")]
    for i, (t, s) in enumerate(pasos):
        c += _caja(_X[i], 130, _W, 52, t, s, _PROCESO_BG, _PROCESO_BR, _PROCESO_TX)
        if i < 2:
            c += _flecha(_X[i] + _W + 2, 156, _X[i + 1] - 6, 156)
    for i in range(3):
        c += _flecha(340, 202, _X[i] + _W / 2, 232)
    salidas = [("Excel consolidado", "descarga", _SALIDA_BG, _SALIDA_BR, _SALIDA_TX),
               ("Word de feedback", "uno por proveedor", _SALIDA_BG, _SALIDA_BR, _SALIDA_TX),
               ("Queda en la sesión", "lo usa la etapa 3", _SESION_BG, _SESION_BR, _SESION_TX)]
    for i, (t, s, bg, br, tx) in enumerate(salidas):
        c += _caja(_X[i], 236, _W, 52, t, s, bg, br, tx)
    return _svg(300, c, "Flujo del módulo 1, consolidador")


def tecnica() -> str:
    c = _caja(230, 10, 220, 52, "Matriz de evaluación", "2 hojas del Excel",
              _ENTRADA_BG, _ENTRADA_BR, _ENTRADA_TX)
    c += _flecha(340, 62, 340, 92)
    c += _contenedor(40, 96, 600, 106)
    pasos = [("Lee la matriz", "7 criterios con peso"),
             ("Puntaje por ítem", "ponderado sobre 5.0"),
             ("Semáforo", "3.75 alto · 2.5 medio")]
    for i, (t, s) in enumerate(pasos):
        c += _caja(_X[i], 130, _W, 52, t, s, _PROCESO_BG, _PROCESO_BR, _PROCESO_TX)
        if i < 2:
            c += _flecha(_X[i] + _W + 2, 156, _X[i + 1] - 6, 156)
    for i in range(3):
        c += _flecha(340, 202, _X[i] + _W / 2, 232)
    salidas = [("Reporte Word", "uno por proveedor", _SALIDA_BG, _SALIDA_BR, _SALIDA_TX),
               ("Excel de puntajes", "respaldo opcional", _SALIDA_BG, _SALIDA_BR, _SALIDA_TX),
               ("Queda en la sesión", "clave ID + proveedor", _SESION_BG, _SESION_BR, _SESION_TX)]
    for i, (t, s, bg, br, tx) in enumerate(salidas):
        c += _caja(_X[i], 236, _W, 52, t, s, bg, br, tx)
    return _svg(300, c, "Flujo del módulo 2, evaluación técnica")


def escenarios() -> str:
    entradas = [("Consolidado", "sesión o archivo", _SESION_BG, _SESION_BR, _SESION_TX),
                ("Puntajes técnicos", "sesión o archivo", _SESION_BG, _SESION_BR, _SESION_TX),
                ("Opcionales", "factor, historial, plan", _ENTRADA_BG, _ENTRADA_BR, _ENTRADA_TX)]
    c = ""
    for i, (t, s, bg, br, tx) in enumerate(entradas):
        c += _caja(_X[i], 10, _W, 52, t, s, bg, br, tx)
    c += _flecha(144, 62, 300, 92)
    c += _flecha(340, 62, 340, 92)
    c += _flecha(536, 62, 380, 92)
    c += _contenedor(40, 96, 600, 166)
    pasos = [("Cruce de claves", "ID + proveedor"),
             ("Factor e historial", "solo si los subes"),
             ("Baseline y ahorro", "fórmulas de Excel")]
    for i, (t, s) in enumerate(pasos):
        c += _caja(_X[i], 130, _W, 52, t, s, _PROCESO_BG, _PROCESO_BR, _PROCESO_TX)
        if i < 2:
            c += _flecha(_X[i] + _W + 2, 156, _X[i + 1] - 6, 156)
    c += _flecha(536, 182, 536, 204)
    c += _caja(252, 208, 372, 38,
               "E1 más barato · E2 mejor puntaje · E3 manual", None,
               _ESC_BG, _ESC_BR, _ESC_TX, size=12)
    c += _flecha(340, 262, 200, 292)
    c += _flecha(340, 262, 480, 292)
    c += _caja(112, 296, _W, 52, "Workbook final", "7 hojas con fórmulas",
               _SALIDA_BG, _SALIDA_BR, _SALIDA_TX)
    c += _caja(392, 296, _W, 52, "Historial de rondas", "descárgalo siempre",
               _SALIDA_BG, _SALIDA_BR, _SALIDA_TX)
    return _svg(360, c, "Flujo del módulo 3, propuesta de adjudicación")
