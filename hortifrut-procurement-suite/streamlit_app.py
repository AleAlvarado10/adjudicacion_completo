"""
Hortifrut Procurement Suite — punto de entrada único.

Reemplaza los tres links de Streamlit que existían por separado:
  1. Consolidador de cotizaciones
  2. Evaluación técnica
  3. Propuesta de adjudicación (escenarios)

Ejecutar en local:  streamlit run streamlit_app.py
"""

import streamlit as st

st.set_page_config(
    page_title="Hortifrut · Procurement Suite",
    page_icon="🫐",
    layout="wide",
    initial_sidebar_state="expanded",
)

from core.shared import estado as E  # noqa: E402
from core.shared import ui  # noqa: E402


# ─── Control de acceso (opcional) ─────────────────────────────────────────────
# Se activa solo si defines una contraseña en los secrets de Streamlit Cloud:
#     app_password = "tu-clave"
# Si no la defines, la app abre directamente.

def _acceso_permitido() -> bool:
    try:
        clave = st.secrets.get("app_password")
    except Exception:
        # Sin archivo de secrets configurado: la app abre sin contraseña.
        clave = None
    if not clave:
        return True
    if st.session_state.get("_auth_ok"):
        return True

    ui.header("Hortifrut · Procurement Suite", "Acceso restringido al equipo de Procurement")
    ingresada = st.text_input("Contraseña del equipo", type="password")
    if ingresada:
        if ingresada == clave:
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


if not _acceso_permitido():
    st.stop()

E.init()

# ─── Navegación ───────────────────────────────────────────────────────────────

paginas = [
    st.Page("app_pages/inicio.py", title="Inicio", icon="🏠", default=True),
    st.Page("app_pages/consolidador.py", title="1. Consolidar cotizaciones", icon="📊"),
    st.Page("app_pages/evaluacion_tecnica.py", title="2. Evaluación técnica", icon="🌿"),
    st.Page("app_pages/escenarios.py", title="3. Propuesta de adjudicación", icon="📈"),
    st.Page("app_pages/historial.py", title="Historial de rondas", icon="🗂️"),
]

st.navigation(paginas).run()
