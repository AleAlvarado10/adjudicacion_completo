"""
Evaluación Técnica — generación del reporte Word por proveedor.

Extraído de `app_evaluacion_hortifruit.py`. La escala de puntajes se recibe como
argumento (`escala`) en lugar de leerse de un global.
"""

import io

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from .parser import SCORE_LABELS, get_feedback, semaforo

def score_rgb(score):
    if score >= 4.0:   return RGBColor(0x1E, 0x84, 0x49)
    elif score >= 3.0: return RGBColor(0xD6, 0x89, 0x10)
    else:              return RGBColor(0xC0, 0x39, 0x2B)

def set_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def heading(doc, text, level, color="1A5276"):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(int(color[:2],16), int(color[2:4],16), int(color[4:],16))
    return h

def generate_report(provider, prov_idx, providers, items, escala=None):
    pname = provider["name"]
    doc   = Document()

    for sec in doc.sections:
        sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
        sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    # Portada
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("HORTIFRUIT")
    r.bold = True; r.font.size = Pt(24)
    r.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Reporte de Evaluación Técnica – Materiales de Riego Importados")
    r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x2E, 0x86, 0xC1)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Proveedor: {pname}")
    r.bold = True; r.font.size = Pt(17)

    doc.add_page_break()

    # Estadísticas
    total_items  = len(items)
    all_scores   = [
        info["totales"][prov_idx]
        for info in items.values()
        if prov_idx < len(info["totales"]) and info["totales"][prov_idx] > 0
    ]
    n_no_cotizo = total_items - len(all_scores)
    avg     = round(sum(all_scores)/len(all_scores), 2) if all_scores else 0.0
    n_alto  = sum(1 for s in all_scores if s >= 3.75)
    n_medio = sum(1 for s in all_scores if 2.5 <= s < 3.75)
    n_bajo  = sum(1 for s in all_scores if s < 2.5)
    sem_i, sem_lbl = semaforo(avg)

    heading(doc, "1. Resumen Ejecutivo", level=1)
    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, (h2, v) in enumerate([
        ("Puntaje Técnico Promedio", f"{avg:.2f} / 5.0"),
        ("Semáforo", f"{sem_i} {sem_lbl}"),
        ("Ítems Cotizados", str(len(all_scores))),
        ("Ítems No Cotizados", str(n_no_cotizo)),
    ]):
        set_bg(tbl.rows[0].cells[j], "1A5276")
        r = tbl.rows[0].cells[j].paragraphs[0].add_run(h2)
        r.bold = True; r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(9)
        tbl.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        bg_val = "EAF2FB" if j < 3 else "FDEBD0"
        set_bg(tbl.rows[1].cells[j], bg_val)
        r2 = tbl.rows[1].cells[j].paragraphs[0].add_run(v)
        r2.bold = True; r2.font.size = Pt(12)
        tbl.rows[1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Distribución por nivel (ítems cotizados): ").bold = True
    p.add_run(f"🟢 Alto ({n_alto})   🟡 Medio ({n_medio})   🔴 Bajo ({n_bajo})   ⬜ No cotizó ({n_no_cotizo})")
    doc.add_paragraph()

    heading(doc, "2. Evaluación Técnica por Ítem", level=1)

    particiones = {}
    for iid, info in items.items():
        particiones.setdefault(info.get("particion","SIN CATEGORÍA"), []).append((iid, info))

    for sec_num, (part, items_list) in enumerate(particiones.items(), 1):
        heading(doc, f"2.{sec_num}. {part}", level=2, color="1F618D")

        for iid, info in items_list:
            score  = info["totales"][prov_idx] if prov_idx < len(info["totales"]) else 0.0
            cotizo = score > 0

            # Ítems no cotizados se omiten aquí — aparecen al final
            if not cotizo:
                continue

            sem_i2, sem_lbl2 = semaforo(score)

            # Título ítem
            p_item = doc.add_paragraph()
            p_item.paragraph_format.space_before = Pt(4)
            p_item.paragraph_format.space_after  = Pt(0)
            ri = p_item.add_run(f"[{iid}] {info['descripcion']}")
            ri.bold = True; ri.font.size = Pt(10)
            ri.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

            # Puntaje
            p_sc = doc.add_paragraph()
            p_sc.paragraph_format.left_indent  = Inches(0.2)
            p_sc.paragraph_format.space_before = Pt(0)
            p_sc.paragraph_format.space_after  = Pt(2)
            p_sc.add_run("Puntaje técnico: ")
            rs = p_sc.add_run(f"{score:.2f} / 5.0")
            rs.bold = True; rs.font.color.rgb = score_rgb(score)
            p_sc.add_run(f"   {sem_i2} {sem_lbl2}")

            # Tabla de criterios
            criterios = info.get("criterios", [])
            if criterios:
                ct = doc.add_table(rows=len(criterios)+1, cols=4)
                ct.style = "Table Grid"
                for j, hdr in enumerate(["Criterio (Peso)", "Puntaje", "Calificación", "Feedback"]):
                    set_bg(ct.rows[0].cells[j], "2471A3")
                    r = ct.rows[0].cells[j].paragraphs[0].add_run(hdr)
                    r.bold = True; r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(8)

                for ri2, crit in enumerate(criterios):
                    score_c = crit["scores"][prov_idx] if prov_idx < len(crit["scores"]) else 0
                    bg      = "EBF5FB" if ri2 % 2 == 0 else "FDFEFE"
                    row_c   = ct.rows[ri2+1]

                    set_bg(row_c.cells[0], bg)
                    row_c.cells[0].paragraphs[0].add_run(f"{crit['nombre']} ({crit['peso']} pts)").font.size = Pt(8)

                    set_bg(row_c.cells[1], bg)
                    p2 = row_c.cells[1].paragraphs[0]
                    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if score_c == 0:
                        r2 = p2.add_run("S/C")
                        r2.italic = True; r2.font.size = Pt(9)
                        r2.font.color.rgb = RGBColor(0x95, 0x96, 0x95)
                    else:
                        r2 = p2.add_run(f"{score_c} / 5")
                        r2.bold = True; r2.font.size = Pt(9); r2.font.color.rgb = score_rgb(score_c)

                    set_bg(row_c.cells[2], bg)
                    p3 = row_c.cells[2].paragraphs[0]
                    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if score_c == 0:
                        r3 = p3.add_run("Sin clasificar")
                        r3.italic = True; r3.font.size = Pt(8)
                        r3.font.color.rgb = RGBColor(0x95, 0x96, 0x95)
                    else:
                        p3.add_run(SCORE_LABELS.get(score_c, "")).font.size = Pt(8)

                    set_bg(row_c.cells[3], bg)
                    if score_c == 0:
                        r4 = row_c.cells[3].paragraphs[0].add_run("No se registró puntaje para este criterio.")
                        r4.italic = True; r4.font.size = Pt(8)
                        r4.font.color.rgb = RGBColor(0x95, 0x96, 0x95)
                    else:
                        fb = get_feedback(escala, crit["nombre"], score_c)
                        row_c.cells[3].paragraphs[0].add_run(fb).font.size = Pt(8)

                # Párrafo separador mínimo tras tabla
                sep = doc.add_paragraph()
                sep.paragraph_format.space_before = Pt(2)
                sep.paragraph_format.space_after  = Pt(2)

    # ── Sección: Ítems No Cotizados ─────────────────────────────────────────
    no_cotizados = [
        (iid, info)
        for iid, info in items.items()
        if prov_idx >= len(info["totales"]) or info["totales"][prov_idx] == 0.0
    ]

    heading(doc, "3. Ítems No Cotizados", level=1)
    p_nc_intro = doc.add_paragraph()
    if no_cotizados:
        r_intro = p_nc_intro.add_run(
            f"{pname} no presentó cotización para {len(no_cotizados)} ítem(s) "
            f"de un total de {len(items)} solicitados."
        )
        r_intro.font.size = Pt(10)

        # Agrupar por partición
        partes_nc = {}
        for iid, info in no_cotizados:
            partes_nc.setdefault(info.get("particion","SIN CATEGORÍA"), []).append((iid, info))

        for part_nc, nc_list in partes_nc.items():
            h_nc = doc.add_paragraph()
            h_nc.paragraph_format.space_before = Pt(4)
            h_nc.paragraph_format.space_after  = Pt(2)
            r_nc = h_nc.add_run(part_nc)
            r_nc.bold = True; r_nc.font.size = Pt(9)
            r_nc.font.color.rgb = RGBColor(0x1F, 0x61, 0x8D)

            nc_tbl = doc.add_table(rows=len(nc_list)+1, cols=3)
            nc_tbl.style = "Table Grid"

            for j, hdr in enumerate(["ID", "Descripción", "Subcategoría"]):
                set_bg(nc_tbl.rows[0].cells[j], "7F8C8D")
                r = nc_tbl.rows[0].cells[j].paragraphs[0].add_run(hdr)
                r.bold = True; r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(8)

            for ri_nc, (iid_nc, info_nc) in enumerate(nc_list):
                row_nc = nc_tbl.rows[ri_nc+1]
                bg_nc = "F2F3F4" if ri_nc % 2 == 0 else "FDFEFE"
                set_bg(row_nc.cells[0], bg_nc)
                row_nc.cells[0].paragraphs[0].add_run(iid_nc).font.size = Pt(8)
                set_bg(row_nc.cells[1], bg_nc)
                row_nc.cells[1].paragraphs[0].add_run(info_nc["descripcion"]).font.size = Pt(8)
                set_bg(row_nc.cells[2], bg_nc)
                row_nc.cells[2].paragraphs[0].add_run(info_nc.get("subcategoria","")).font.size = Pt(8)

            sep2 = doc.add_paragraph()
            sep2.paragraph_format.space_before = Pt(2)
            sep2.paragraph_format.space_after  = Pt(2)
    else:
        r_intro = p_nc_intro.add_run("✅ El proveedor cotizó todos los ítems solicitados.")
        r_intro.font.size = Pt(10)
        r_intro.font.color.rgb = RGBColor(0x1E, 0x84, 0x49)

    heading(doc, "4. Conclusión Técnica", level=1)
    if avg >= 3.75:
        txt = (f"{pname} alcanza un puntaje técnico promedio de {avg:.2f}/5.0 (Nivel ALTO). "
               f"Se destacó con {n_alto} ítems en nivel alto de un total de {len(all_scores)} evaluados. "
               "Se recomienda proceder con la adjudicación de los ítems donde obtiene nivel ALTO.")
    elif avg >= 2.5:
        txt = (f"{pname} alcanza un puntaje técnico promedio de {avg:.2f}/5.0 (Nivel MEDIO). "
               f"Presenta fortalezas en {n_alto} ítems con nivel alto, con oportunidades de mejora "
               f"en {n_bajo} ítems con nivel bajo. "
               "Se recomienda solicitar mejoras antes de adjudicar.")
    else:
        txt = (f"{pname} alcanza un puntaje técnico promedio de {avg:.2f}/5.0 (Nivel BAJO), "
               "inferior al mínimo recomendado de 3.0/5.0. "
               "No se recomienda adjudicar sin mejoras sustanciales.")
    doc.add_paragraph(txt)

    pf = doc.add_paragraph()
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = pf.add_run("Documento generado automáticamente – Sistema de Evaluación Técnica HORTIFRUIT")
    rf.italic = True; rf.font.size = Pt(8); rf.font.color.rgb = RGBColor(0x7F,0x8C,0x8D)

    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf.getvalue()

