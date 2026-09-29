import streamlit as st

from core.shared import diagrama
from core.shared import estado as E
from core.shared import ui

ui.header(
    "Hortifrut · Procurement Suite",
    "Consolidación económica · Evaluación técnica · Propuesta de adjudicación",
)

nombre, ronda = ui.barra_proceso()

st.markdown(
    "Las tres herramientas que antes estaban en links separados ahora son un solo "
    "flujo. Lo que produce una etapa queda disponible para la siguiente **sin "
    "descargar ni volver a subir archivos**."
)

# ─── Estado del pipeline ──────────────────────────────────────────────────────

e1, e2, e3 = E.etapa1_lista(), E.etapa2_lista(), E.etapa3_lista()


def _estado(ok, texto_ok, texto_pendiente):
    return f"✅ {texto_ok}" if ok else f"⚪ {texto_pendiente}"


c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("#### 1. Consolidar cotizaciones")
    st.caption("Entra: los Excel de cada proveedor. Sale: el consolidado comparativo "
               "y el Word de feedback económico.")
    if e1:
        df = E.etapa(1)["df_wide"]
        n_prov = len([c for c in df.columns if c.startswith("PU_Total__")])
        st.success(f"{len(df)} ítems · {n_prov} proveedores en memoria")
    else:
        st.info("Pendiente")
    ui.link_pagina("app_pages/consolidador.py", "Ir a la etapa 1 →")

with c2:
    st.markdown("#### 2. Evaluación técnica")
    st.caption("Entra: la matriz de evaluación. Sale: el reporte Word por proveedor "
               "y los puntajes /5.0 para la etapa 3.")
    if e2:
        st.success(f"{len(E.etapa(2)['puntajes'])} puntajes disponibles")
    else:
        st.info("Pendiente (opcional)")
    ui.link_pagina("app_pages/evaluacion_tecnica.py", "Ir a la etapa 2 →")

with c3:
    st.markdown("#### 3. Propuesta de adjudicación")
    st.caption("Entra: el consolidado (y los puntajes, si existen). Sale: el workbook "
               "de 7 hojas con escenarios y ahorro.")
    if e3:
        st.success("Workbook generado en esta sesión")
    elif e1:
        st.info("Listo para ejecutarse con los datos de la etapa 1")
    else:
        st.info("Pendiente")
    ui.link_pagina("app_pages/escenarios.py", "Ir a la etapa 3 →")

st.markdown("---")

st.markdown("### Cómo funciona cada módulo")
st.caption(
    "Cada etapa recibe archivos, hace su proceso y deja resultados descargables. "
    "Lo que queda en la sesión es lo que evita volver a subir archivos en la etapa siguiente."
)

tab1, tab2, tab3 = st.tabs(
    ["1. Consolidar cotizaciones", "2. Evaluación técnica", "3. Propuesta de adjudicación"]
)

with tab1:
    st.markdown(diagrama.consolidador(), unsafe_allow_html=True)
    st.caption(
        "Detecta la hoja de cotización y la razón social aunque la plantilla cambie de "
        "formato, descarta los ítems no cotizados y compara a todos los proveedores "
        "lado a lado. El Word de feedback ya no necesita Node.js."
    )

with tab2:
    st.markdown(diagrama.tecnica(), unsafe_allow_html=True)
    st.caption(
        "Independiente de la etapa 1: puede correrse antes o después. Los puntajes "
        "quedan guardados con la clave ID + proveedor, que es la que usa la etapa 3 "
        "para cruzarlos."
    )

with tab3:
    st.markdown(diagrama.escenarios(), unsafe_allow_html=True)
    st.caption(
        "Es el único módulo que depende de los otros dos. Si no subes la tabla de "
        "factor logístico, el archivo se genera sin las columnas de factor. El "
        "historial de rondas se acumula por número de ronda."
    )

st.markdown("---")

st.markdown(
    f"""
**Estado actual del proceso**

Etapa 1 {_estado(e1, 'lista', 'pendiente')} ·
etapa 2 {_estado(e2, 'lista', 'pendiente')} ·
etapa 3 {_estado(e3, 'lista', 'pendiente')}
"""
)

st.warning(
    "El historial de rondas y los datos cargados viven **solo en esta sesión**. "
    "Streamlit Cloud reinicia el contenedor periódicamente: descarga siempre los "
    "archivos que necesites conservar antes de cerrar.",
    icon="⚠️",
)

st.markdown("---")
if st.button("🔄 Reiniciar el proceso completo"):
    E.reiniciar()
    st.rerun()
