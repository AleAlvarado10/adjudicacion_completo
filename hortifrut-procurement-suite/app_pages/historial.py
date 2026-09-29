import io

import pandas as pd
import streamlit as st

from core.escenarios.rondas import cargar_historial, historial_wide
from core.shared.archivos import buffer
from core.shared import estado as E
from core.shared import ui

ui.header("🗂️ Historial de rondas", "Comparación de precios entre rondas de negociación")
ui.barra_proceso()
ui.flujo(
    "<b>Entra:</b> el historial de esta sesión o un archivo de una ronda previa. "
    "&nbsp;·&nbsp; <b>Sale:</b> la comparación ronda a ronda. &nbsp;·&nbsp; "
    "<b>Depende de:</b> etapa 3."
)

st.warning(
    "Streamlit Cloud no tiene disco persistente: el historial vive en la sesión y se "
    "pierde al reiniciarse el contenedor. **Descárgalo al terminar cada ronda** y "
    "vuelve a subirlo en la siguiente.",
    icon="⚠️",
)

historial = E.etapa(3)["historial"]
if historial is None:
    historial = st.session_state.get("historial_df")

subido = st.file_uploader("Cargar un historial guardado (.xlsx)", type=["xlsx"])
if subido is not None:
    try:
        historial = cargar_historial(buffer(subido))
        st.session_state["historial_df"] = historial
        st.success("Historial cargado.", icon="✅")
    except Exception as e:
        st.error(f"No se pudo leer el historial: {e}")

if historial is None or historial.empty:
    st.info(
        "Todavía no hay historial. Corre la etapa 3 para generarlo, o sube un archivo "
        "de una ronda anterior."
    )
    ui.link_pagina("app_pages/escenarios.py", "Ir a la etapa 3 →")
    st.stop()

st.markdown("### Comparación entre rondas")
wide = historial_wide(historial)
if not wide.empty:
    st.dataframe(wide, use_container_width=True, height=420)
else:
    st.caption("Aún no hay suficientes rondas para comparar.")

with st.expander("Ver el historial completo (formato largo)"):
    st.dataframe(historial, use_container_width=True, height=360)

c1, c2 = st.columns(2)
with c1:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        historial.to_excel(w, sheet_name="Historial", index=False)
    buf.seek(0)
    st.download_button(
        "⬇️ Descargar historial (.xlsx)", buf.getvalue(), "historial_rondas.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    n_rondas = historial["Ronda"].nunique() if "Ronda" in historial.columns else 0
    ui.tarjeta_metrica(n_rondas, "rondas registradas")
