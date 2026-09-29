"""
Evaluación Técnica — lectura de la Matriz Excel.

Extraído de `app_evaluacion_hortifruit.py` sin cambios de lógica. Única diferencia:
la escala de puntajes ya NO vive en un `global` (`CRITERION_FEEDBACK`), sino que se
devuelve junto con los datos y se pasa como argumento. En Streamlit Cloud varios
usuarios comparten el mismo proceso y un global se filtraría entre sesiones.
"""

import io
import re
import openpyxl

SCORE_LABELS = {5: "Excelente", 4: "Bueno", 3: "Aceptable", 2: "Deficiente", 1: "Muy Deficiente"}


def get_feedback(escala, nombre, score):
    """Busca el feedback en la escala cargada desde el Excel."""
    nombre_up = nombre.upper()
    for key, fb in (escala or {}).items():
        key_words = key.upper().split()[:3]
        if sum(1 for w in key_words if w in nombre_up) >= 2:
            return fb.get(score, f"Puntaje {score}/5 – {SCORE_LABELS.get(score, 'N/D')}")
    return f"Puntaje {score}/5 – {SCORE_LABELS.get(score, 'N/D')}"


def semaforo(score):
    if score >= 3.75:  return "🟢", "ALTO"
    elif score >= 2.5: return "🟡", "MEDIO"
    else:              return "🔴", "BAJO"

def load_escala_puntajes(wb):
    """
    Lee la hoja ESCALA DE PUNTAJES.
    Estructura: col B=nombre criterio (solo en primera fila del criterio),
                col C=puntaje (1-5), col D=calificación, col E=descripción.
    """
    escala = {}
    try:
        ws = wb["ESCALA DE PUNTAJES"]
        rows = list(ws.iter_rows(values_only=True))
        current_criterio = None
        for row in rows:
            col_b = str(row[1]).strip() if row[1] else ""
            col_c = row[2]   # puntaje numérico
            col_e = str(row[4]).strip() if row[4] else ""
            # Fila con nombre de criterio (col B tiene texto, col C tiene número)
            if col_b and col_b not in ("CRITERIO", "ESCALA DE PUNTAJES POR CRITERIO"):
                try:
                    puntaje = int(col_c)
                    current_criterio = col_b
                    escala[current_criterio] = {}
                    if col_e:
                        escala[current_criterio][puntaje] = col_e
                except (ValueError, TypeError):
                    pass
            # Fila de puntaje adicional (col B vacía, col C tiene número)
            elif col_b == "" and col_c is not None and current_criterio:
                try:
                    puntaje = int(col_c)
                    if col_e:
                        escala[current_criterio][puntaje] = col_e
                except (ValueError, TypeError):
                    pass
    except Exception:
        pass
    return escala


def parse_matrix(file_bytes):
    """Devuelve (providers, items, escala)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    # La escala de puntajes se devuelve, no se guarda en un global.
    escala = load_escala_puntajes(wb)
    ws = wb["EVALUACION POR ITEM"]
    rows = list(ws.iter_rows(values_only=True))

    # ── Detectar proveedores y sus columnas desde fila 3 (idx 2) ─────────────
    # Patrón: col 4=P1_nombre, 5=P1_puntaje_header, col 6=P2_nombre, ...
    # En fila 4 (idx 3): col 4=PUNTAJE, col 5=POND.AUTO, col 6=PUNTAJE, col 7=POND...
    header_names = rows[2]   # fila 3: nombres
    header_cols  = rows[3]   # fila 4: PUNTAJE / POND.AUTO

    providers = []   # lista de {name, col_score, col_pond}
    col = 4
    while col < len(header_names):
        name = header_names[col]
        if name and str(name).strip():
            name_clean = str(name).strip()
            # columna de puntaje = col, columna de pond = col+1
            providers.append({
                "name":      name_clean,
                "col_score": col,
                "col_pond":  col + 1,
            })
            col += 2
        else:
            col += 1

    if not providers:
        raise ValueError("No se encontraron proveedores en la fila 3 del Excel.")

    n = len(providers)

    # ── Parsear ítems ─────────────────────────────────────────────────────────
    items = {}
    current_item = None
    current_part = ""

    for row in rows:
        val0 = str(row[0]).strip() if row[0] else ""

        # Detectar partición (no tiene RIEGxxx ni PUNTAJE ÍTEM)
        upper = val0.upper()
        if val0 and not re.search(r"RIEG\d+", val0) and "PUNTAJE" not in val0 \
                and "HORTIFRUT" not in val0 and "Ingrese" not in val0 and val0 != "ID":
            if "RED HIDRAÚLICA" in upper or ("RED" in upper and "HIDRA" in upper):
                current_part = "RED HIDRAÚLICA DE RIEGO"
            elif "EMISOR" in upper:
                current_part = "EMISORES DE RIEGO"
            elif "AUTOMATIZ" in upper:
                current_part = "AUTOMATIZACIÓN Y ELECTRIFICACIÓN"
            elif "CABEZAL" in upper:
                current_part = "CABEZAL DE RIEGO"

        # Fila de ítem: col0 contiene RIEGxxx
        m = re.search(r"(RIEG\d+)", val0)
        if m and "PUNTAJE" not in val0:
            current_item = m.group(1)
            items[current_item] = {
                "descripcion":  str(row[1]).strip() if row[1] else "",
                "subcategoria": str(row[2]).strip() if row[2] else "",
                "particion":    current_part,
                "criterios":    [],
                "totales":      [0.0] * n,
            }

        # Fila de criterio: col3 tiene "(N pts)"
        elif row[0] is None and row[3] and "(" in str(row[3]):
            if current_item:
                val3  = str(row[3]).strip()
                nombre = re.sub(r'\s*\(\d+\s*pts?\)', '', val3, flags=re.IGNORECASE).strip()
                peso_m = re.search(r'\((\d+)\s*pts?\)', val3, re.IGNORECASE)
                peso   = int(peso_m.group(1)) if peso_m else 0
                scores = []
                for p in providers:
                    ci = p["col_score"]
                    v  = row[ci] if ci < len(row) else None
                    try:    scores.append(int(v) if v is not None else 0)
                    except: scores.append(0)
                items[current_item]["criterios"].append({
                    "nombre": nombre, "peso": peso, "scores": scores
                })

        # Fila de total: col0 tiene "PUNTAJE ÍTEM"
        elif "PUNTAJE ÍTEM" in val0 and current_item:
            tots = []
            for p in providers:
                ci = p["col_pond"]
                v  = row[ci] if ci < len(row) else None
                try:    tots.append(round(float(v), 2) if v is not None else 0.0)
                except: tots.append(0.0)
            items[current_item]["totales"] = tots

    return providers, items, escala
