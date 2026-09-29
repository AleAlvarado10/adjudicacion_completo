"""
Etapa 2 — Evaluación técnica.
Adaptado de `app_evaluacion_hortifruit.py`. Cambios respecto al original:
  - la lógica vive en core/tecnica/ (parser + reporte_word), sin globals
  - los puntajes por ítem quedan en session_state con la llave "ID|Proveedor",
    que es exactamente lo que la etapa 3 necesita
"""

import io
import re
import zipfile

import pandas as pd
import streamlit as st

from core.shared.archivos import bytes_de
from core.shared import estado as E
from core.shared import ui
from core.tecnica.parser import parse_matrix, semaforo
from core.tecnica.reporte_word import generate_report

ui.header(
    "🌿 Etapa 2 · Evaluación técnica de materiales",
    "Reportes Word por proveedor · solo criterios técnicos",
)
nombre_proceso, ronda = ui.barra_proceso()
ui.flujo(
    "<b>Entra:</b> la matriz de evaluación (hojas <i>EVALUACION POR ITEM</i> y "
    "<i>ESCALA DE PUNTAJES</i>). &nbsp;·&nbsp; <b>Sale:</b> el reporte Word por "
    "proveedor y los puntajes /5.0 para la etapa 3. &nbsp;·&nbsp; "
    "<b>Depende de:</b> nada (puede correrse antes o después de la etapa 1)."
)

with st.sidebar:
    st.markdown("### ⚙️ Matriz de evaluación")
    uploaded = st.file_uploader("Cargar Matriz Excel (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.markdown(
        """
**Semáforo técnico:**
- 🟢 ALTO ≥ 3.75
- 🟡 MEDIO 2.5–3.74
- 🔴 BAJO < 2.5

**Criterios técnicos (7):**
- Especif. Técnica (25 pts)
- Marca Requerida (5 pts)
- Tiempo de Entrega (8 pts)
- País de Origen (7 pts)
- Garantía (5 pts)
- Soporte Técnico (5 pts)
- Certificaciones (5 pts)
"""
    )

# ─── Carga ────────────────────────────────────────────────────────────────────

if uploaded is None:
    if E.etapa2_lista():
        st.success(
            f"Ya hay {len(E.etapa(2)['puntajes'])} puntajes técnicos en esta sesión, "
            "disponibles para la etapa 3. Sube una matriz nueva para reemplazarlos.",
            icon="✅",
        )
        ui.link_pagina("app_pages/escenarios.py", "Ir a la etapa 3 →")
    else:
        st.info("👈 Carga el archivo Excel de la Matriz de Evaluación para comenzar.")
    st.stop()

with st.spinner("Procesando la matriz..."):
    try:
        file_bytes = bytes_de(uploaded)
        providers, items, escala = parse_matrix(file_bytes)
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

# Proveedores realmente evaluados vs. columnas vacías de la plantilla
def _n_items_con_puntaje(idx):
    return sum(1 for inf in items.values()
               if idx < len(inf["totales"]) and inf["totales"][idx] > 0)

evaluados = [p for i, p in enumerate(providers) if _n_items_con_puntaje(i) > 0]
sin_evaluar = [p["name"] for i, p in enumerate(providers) if _n_items_con_puntaje(i) == 0]

st.sidebar.success(f"✅ {len(items)} ítems · {len(evaluados)} proveedores con puntajes")
st.sidebar.markdown("**Proveedores evaluados:**")
for p in evaluados:
    st.sidebar.markdown(f"- {p['name']} ({_n_items_con_puntaje(providers.index(p))} ítems)")
if sin_evaluar:
    st.sidebar.caption("Sin puntajes (columnas vacías): " + ", ".join(sin_evaluar))

if not evaluados:
    st.error(
        "La matriz se leyó bien, pero **ningún proveedor tiene puntajes cargados**. "
        "Si los puntajes de tu plantilla se calculan con fórmulas, abre el archivo en "
        "Excel, deja que recalcule y vuelve a guardarlo antes de subirlo: la app lee "
        "el último valor guardado, no recalcula fórmulas.",
        icon="🚨",
    )
    st.stop()

# ─── Guardar puntajes para la etapa 3 ─────────────────────────────────────────
# La etapa 3 cruza por la llave "ID|Proveedor" (ver core/escenarios/evaluacion.py).
puntajes = {}
for iid, info in items.items():
    for i, prov in enumerate(providers):
        if i < len(info["totales"]) and info["totales"][i] > 0:
            puntajes[f"{iid}|{prov['name']}"] = float(info["totales"][i])

E.guardar_etapa(2, providers=providers, items=items, escala=escala, puntajes=puntajes)

# ─── Vista previa ─────────────────────────────────────────────────────────────

st.subheader("📊 Puntajes técnicos por proveedor")
n_cols = min(len(evaluados), 5)
cols = st.columns(n_cols)
for i, prov in enumerate(providers):
    if prov not in evaluados:
        continue
    sc = [
        info["totales"][i]
        for info in items.values()
        if i < len(info["totales"]) and info["totales"][i] > 0
    ]
    avg = round(sum(sc) / len(sc), 2) if sc else 0.0
    sem_i, _ = semaforo(avg)
    with cols[evaluados.index(prov) % n_cols]:
        st.metric(f"{sem_i} {prov['name']}", f"{avg:.2f} / 5.0", f"{len(sc)} ítems")

st.divider()

particiones_disp = {}
for iid, info in items.items():
    particiones_disp.setdefault(info.get("particion", "SIN CAT"), []).append((iid, info))

for part, items_list in particiones_disp.items():
    with st.expander(f"📂 {part} ({len(items_list)} ítems)", expanded=False):
        table_rows = []
        for iid, info in items_list:
            row = {"ID": iid, "Descripción": info["descripcion"][:40]}
            for i, prov in enumerate(providers):
                s = info["totales"][i] if i < len(info["totales"]) else 0.0
                sem_p, _ = semaforo(s)
                row[prov["name"][:12]] = f"{sem_p} {s:.2f}" if s > 0 else "—"
            table_rows.append(row)
        st.dataframe(table_rows, use_container_width=True, height=280)

st.divider()

# ─── Reportes Word ────────────────────────────────────────────────────────────

st.subheader("📄 Generar reportes Word")
prov_names = [p["name"] for p in evaluados]
col1, col2 = st.columns([2, 1])
with col1:
    todos = [p["name"] for p in providers]
    sel_names = st.multiselect(
        "Proveedores a generar", todos, default=prov_names,
        help="Por defecto solo los que tienen puntajes. Puedes agregar los demás si "
             "quieres su reporte igual.",
    )
with col2:
    hacer_zip = st.checkbox("Comprimir en ZIP", value=True)

if st.button("🚀 Generar reportes", type="primary", use_container_width=True):
    if not sel_names:
        st.warning("Selecciona al menos un proveedor.")
    else:
        docs = {}
        bar = st.progress(0)
        for i, pname in enumerate(sel_names):
            prov = next(p for p in providers if p["name"] == pname)
            idx = providers.index(prov)
            docs[pname] = generate_report(prov, idx, providers, items, escala=escala)
            bar.progress((i + 1) / len(sel_names), text=f"✓ {pname}")
        bar.empty()
        st.success(f"✅ {len(docs)} reporte(s) generado(s).")

        if hacer_zip and len(docs) > 1:
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for pname, db in docs.items():
                    safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", pname)
                    zf.writestr(f"Reporte_{safe}_HORTIFRUT.docx", db)
            zbuf.seek(0)
            st.download_button(
                "⬇️ Descargar todos (ZIP)", zbuf.getvalue(),
                "Reportes_Tecnicos_HORTIFRUT.zip", "application/zip",
                use_container_width=True,
            )

        st.markdown("### Descarga individual")
        dl_cols = st.columns(min(len(docs), 3))
        for i, (pname, db) in enumerate(docs.items()):
            with dl_cols[i % 3]:
                safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", pname)
                st.download_button(
                    f"⬇️ {pname}", db,
                    f"Reporte_{safe}_HORTIFRUT.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key=f"dl_{pname}",
                )

# ─── Puntajes disponibles para la etapa 3 ─────────────────────────────────────

st.divider()
st.markdown("#### 🔗 Puntajes para la etapa 3")
st.caption(
    "La etapa 3 cruza los puntajes por la llave `ID | Proveedor`. Estos quedaron "
    "guardados en la sesión; si los IDs o los nombres de proveedor difieren de los "
    "del consolidado, la etapa 3 te avisará cuántos no cruzaron."
)

df_puntajes = pd.DataFrame(
    [{"ID": k.split("|")[0], "Proveedor": k.split("|")[1], "Puntaje": v}
     for k, v in puntajes.items()]
)
st.dataframe(df_puntajes, use_container_width=True, height=240)

buf_p = io.BytesIO()
with pd.ExcelWriter(buf_p, engine="openpyxl") as w:
    df_puntajes.to_excel(w, sheet_name="Puntajes", index=False)
buf_p.seek(0)

col_a, col_b = st.columns(2)
with col_a:
    st.download_button(
        "⬇️ Descargar puntajes (.xlsx)", buf_p.getvalue(),
        f"Puntajes_Tecnicos_{nombre_proceso}_R{ronda}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_b:
    ui.link_pagina("app_pages/escenarios.py", "Continuar a la etapa 3 →")
