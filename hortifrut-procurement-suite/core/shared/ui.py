"""Componentes de interfaz comunes a todas las páginas."""

import streamlit as st

from . import constantes as C

_CSS = f"""
<style>
    html, body, [class*="css"] {{
        font-family: "Century Gothic", Poppins, Futura, sans-serif;
    }}
    .hf-header {{
        background: linear-gradient(135deg, {C.AZUL} 0%, {C.AZUL_CLARO} 100%);
        padding: 1.3rem 1.8rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.2rem;
    }}
    .hf-header h1 {{ margin: 0; font-size: 1.55rem; font-weight: 600; }}
    .hf-header p  {{ margin: 0.35rem 0 0; opacity: 0.9; font-size: 0.9rem; }}

    .hf-flujo {{
        background: {C.MARFIL};
        border-left: 4px solid {C.AZUL_CLARO};
        border-radius: 6px;
        padding: 0.65rem 1rem;
        margin-bottom: 1.2rem;
        font-size: 0.87rem;
        color: {C.GRIS_TEXTO};
    }}

    .metric-card {{
        background: white;
        border: 1px solid #E4E4E4;
        border-radius: 10px;
        padding: 0.9rem 1.1rem;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }}
    .metric-card .value {{ font-size: 1.9rem; font-weight: 700; color: {C.AZUL}; }}
    .metric-card .label {{ font-size: 0.78rem; color: #666; margin-top: 0.2rem; }}

    .prov-card {{
        background: #F2F4FD;
        border-left: 4px solid {C.AZUL_CLARO};
        border-radius: 6px;
        padding: 0.7rem 1rem;
        margin-bottom: 0.5rem;
    }}
    .prov-card.error   {{ background: #FFF0F0; border-left-color: #C00000; }}
    .prov-card.warning {{ background: #FFFBE6; border-left-color: {C.VERDE}; }}

    .step-label {{
        font-weight: 600;
        color: {C.AZUL};
        font-size: 1.03rem;
        margin-bottom: 0.4rem;
    }}

    div[data-testid="stFileUploader"] {{
        border: 2px dashed {C.AZUL_CLARO};
        border-radius: 10px;
        padding: 0.5rem;
    }}
    .stDownloadButton > button {{
        background: linear-gradient(135deg, {C.AZUL}, {C.AZUL_CLARO});
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.65rem 1.8rem;
        font-weight: 600;
        width: 100%;
    }}
</style>
"""


def estilos():
    st.markdown(_CSS, unsafe_allow_html=True)


def header(titulo, subtitulo=""):
    estilos()
    st.markdown(
        f"""<div class="hf-header"><h1>{titulo}</h1><p>{subtitulo}</p></div>""",
        unsafe_allow_html=True,
    )


def flujo(texto):
    """Línea que explica qué entra, qué sale y de qué etapa depende."""
    st.markdown(f'<div class="hf-flujo">{texto}</div>', unsafe_allow_html=True)


def paso(texto):
    st.markdown(f'<div class="step-label">{texto}</div>', unsafe_allow_html=True)


def tarjeta_metrica(valor, etiqueta, size="1.9rem"):
    st.markdown(
        f'<div class="metric-card"><div class="value" style="font-size:{size}">{valor}'
        f'</div><div class="label">{etiqueta}</div></div>',
        unsafe_allow_html=True,
    )


def barra_proceso():
    """Selector de proceso y ronda, compartido por todas las etapas."""
    from . import estado as E

    p = E.proceso()
    with st.sidebar:
        st.markdown("### ⚙️ Proceso")
        nombre = st.text_input(
            "Nombre del proceso", value=p["nombre"],
            help="Se comparte entre las tres etapas. Ej: PROC_001_Colombia_Riego",
        )
        ronda = st.number_input(
            "Número de ronda", min_value=1, max_value=99, value=int(p["ronda"]), step=1,
        )
        E.set_proceso(nombre, ronda)
        st.caption(f"Trabajando en **{nombre}** · Ronda **{int(ronda)}**")
        st.markdown("---")
    return nombre, int(ronda)


def link_pagina(ruta, etiqueta):
    """
    Enlace a otra página de la suite. Envuelto en try/except porque st.page_link
    falla si la página se ejecuta fuera de la navegación (por ejemplo en tests).
    """
    try:
        st.page_link(ruta, label=etiqueta)
    except Exception:
        st.caption(f"➡️ {etiqueta}")
