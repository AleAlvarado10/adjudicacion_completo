"""
Generador de Feedback Económico (Word) por proveedor.

Port 1:1 de `generar_feedback_word.js` (librería `docx` de npm) a python-docx,
para poder correr en Streamlit Community Cloud, donde no existe Node.js.

Se mantienen exactamente las mismas medidas del original:
  - página Letter (12240 x 15840 twips), márgenes 1080 twips (0.75")
  - fuente Arial; tamaños en half-points del JS convertidos a puntos
  - tabla de tarifas de 5 columnas: [950, 2700, 2406, 2000, 1300] twips
  - tabla de leyenda de 3 columnas: [2339, 3508, 3509] twips
  - bordes SINGLE size 4 (0.5 pt) color 808080; márgenes de celda 60/60/80/80
"""

import io

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

FONT = "Arial"

# Colores de semaforización (idénticos al JS original)
COLOR_ROJO = "FF0000"
COLOR_AMARILLO = "FFFF00"
COLOR_VERDE = "00FF00"
COLOR_AZUL = "4C94D8"
GRIS_HEADER = "D9D9D9"

# Anchos en twips (DXA), tal como el original
COLS = [950, 2700, 2406, 2000, 1300]
COLS_LEY = [2339, 3508, 3509]
ANCHO_TOTAL = 9356

FILAS_LEYENDA = [
    {"color": COLOR_AZUL, "min": "1%", "max": "5%"},
    {"color": COLOR_VERDE, "min": "6%", "max": "25%"},
    {"color": COLOR_AMARILLO, "min": "26%", "max": "50%"},
    {"color": COLOR_ROJO, "min": "51%", "max": "Más / Up"},
]


def color_por_feedback(feedback: str) -> str:
    f = (feedback or "").lower()
    if "rojo" in f:
        return COLOR_ROJO
    if "amarillo" in f:
        return COLOR_AMARILLO
    if "verde" in f:
        return COLOR_VERDE
    if "azul" in f:
        return COLOR_AZUL
    return "FFFFFF"


# ─── Helpers de XML (python-docx no expone estas propiedades) ─────────────────

def _sombrear(celda, hex_color):
    if not hex_color:
        return
    tc_pr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _margenes_celda(celda, top=60, bottom=60, left=80, right=80):
    tc_pr = celda._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for lado, valor in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), str(valor))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tc_pr.append(mar)


def _valign_center(celda):
    tc_pr = celda._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tc_pr.append(va)


def _bordes_tabla(tabla, size=4, color="808080"):
    tbl_pr = tabla._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.append(borders)


def _escribir_celda(celda, texto, ancho, bold=False, fill=None,
                    align=WD_ALIGN_PARAGRAPH.CENTER, size=9):
    celda.width = Twips(ancho)
    _sombrear(celda, fill)
    _margenes_celda(celda)
    _valign_center(celda)
    parrafo = celda.paragraphs[0]
    parrafo.alignment = align
    parrafo.paragraph_format.space_after = Pt(0)
    run = parrafo.add_run("" if texto is None else str(texto))
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)


def _titulo_centrado(doc, texto, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(texto)
    run.bold = True
    run.font.name = FONT
    run.font.size = Pt(size)
    return p


# ─── Documento ────────────────────────────────────────────────────────────────

def generar_docx(nombre_proceso: str, ronda, proveedor: str, items: list[dict]) -> bytes:
    """
    Construye el Word de feedback económico y lo devuelve como bytes.
    `items` es la salida de `preparar_items_para_word`.
    """
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), FONT)

    for sec in doc.sections:
        sec.page_width = Twips(12240)
        sec.page_height = Twips(15840)
        sec.top_margin = Twips(1080)
        sec.bottom_margin = Twips(1080)
        sec.left_margin = Twips(1080)
        sec.right_margin = Twips(1080)

    # ── Encabezado ────────────────────────────────────────────────────────────
    _titulo_centrado(doc, f"FEEDBACK ECONÓMICO RONDA {ronda}")
    _titulo_centrado(doc, f"PROCESO {nombre_proceso}")
    _titulo_centrado(doc, str(proveedor).upper(), size=12)

    vacio = doc.add_paragraph()
    vacio.paragraph_format.space_after = Pt(5)

    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.paragraph_format.space_after = Pt(10)
    run = p_intro.add_run(
        "Presentamos a continuación el feedback económico de las tarifas enviadas "
        "mediante un esquema de semaforización, el cual explicamos al final del documento:"
    )
    run.font.name = FONT
    run.font.size = Pt(10)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(5)
    run = p_sub.add_run("Tarifas por Ítem")
    run.bold = True
    run.font.name = FONT
    run.font.size = Pt(11)

    # ── Tabla de tarifas ──────────────────────────────────────────────────────
    tabla = doc.add_table(rows=1 + len(items), cols=5)
    tabla.alignment = WD_TABLE_ALIGNMENT.LEFT
    tabla.autofit = False
    _bordes_tabla(tabla)

    headers = ["ID Ítem", "Descripción", "Descripción del Material Ofertado",
               "P.U. EXW (USD)", "Feedback"]
    for j, h in enumerate(headers):
        _escribir_celda(tabla.rows[0].cells[j], h, COLS[j], bold=True, fill=GRIS_HEADER)

    for i, item in enumerate(items, start=1):
        fila = tabla.rows[i]
        _escribir_celda(fila.cells[0], item.get("id", ""), COLS[0])
        _escribir_celda(fila.cells[1], item.get("descripcion", ""), COLS[1],
                        align=WD_ALIGN_PARAGRAPH.LEFT)
        _escribir_celda(fila.cells[2], item.get("marca_ofertada") or "—", COLS[2],
                        align=WD_ALIGN_PARAGRAPH.LEFT)
        _escribir_celda(fila.cells[3], f"{float(item.get('pu_exw', 0)):.2f}", COLS[3])
        feedback = item.get("feedback", "")
        _escribir_celda(fila.cells[4], feedback, COLS[4], bold=True,
                        fill=color_por_feedback(feedback))

    # ── Leyenda ───────────────────────────────────────────────────────────────
    p_ley = doc.add_paragraph()
    p_ley.paragraph_format.space_before = Pt(15)
    p_ley.paragraph_format.space_after = Pt(5)
    run = p_ley.add_run("Leyenda:")
    run.bold = True
    run.font.name = FONT
    run.font.size = Pt(11)

    tabla_ley = doc.add_table(rows=2 + len(FILAS_LEYENDA), cols=3)
    tabla_ley.alignment = WD_TABLE_ALIGNMENT.LEFT
    tabla_ley.autofit = False
    _bordes_tabla(tabla_ley)

    # Fila 1: "Color" + "Dispersión / Dispersion" abarcando las dos últimas columnas
    _escribir_celda(tabla_ley.rows[0].cells[0], "Color", COLS_LEY[0],
                    bold=True, fill=GRIS_HEADER)
    combinada = tabla_ley.rows[0].cells[1].merge(tabla_ley.rows[0].cells[2])
    _escribir_celda(combinada, "Dispersión / Dispersion", COLS_LEY[1] + COLS_LEY[2],
                    bold=True, fill=GRIS_HEADER)

    # Fila 2: vacío | Min. | Max.
    _escribir_celda(tabla_ley.rows[1].cells[0], "", COLS_LEY[0])
    _escribir_celda(tabla_ley.rows[1].cells[1], "Min.", COLS_LEY[1], bold=True)
    _escribir_celda(tabla_ley.rows[1].cells[2], "Max.", COLS_LEY[2], bold=True)

    # Filas de color (la primera celda va sin texto, solo con relleno — igual que el JS)
    for i, l in enumerate(FILAS_LEYENDA, start=2):
        fila = tabla_ley.rows[i]
        _escribir_celda(fila.cells[0], "", COLS_LEY[0], fill=l["color"])
        _escribir_celda(fila.cells[1], l["min"], COLS_LEY[1])
        _escribir_celda(fila.cells[2], l["max"], COLS_LEY[2])

    # ── Nota final ────────────────────────────────────────────────────────────
    p_nota = doc.add_paragraph()
    p_nota.paragraph_format.space_before = Pt(10)
    run = p_nota.add_run(
        "* El porcentaje de dispersión corresponde a la diferencia entre el precio "
        "total ofertado y el precio mínimo de mercado para cada ítem."
    )
    run.italic = True
    run.font.name = FONT
    run.font.size = Pt(8)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()
