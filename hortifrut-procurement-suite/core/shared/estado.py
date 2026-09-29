"""
Estado compartido de la sesión.

Todo vive bajo `st.session_state["hf"]` para no chocar con otras claves y para
poder reiniciar el proceso completo de una sola vez. Es lo que permite que la
etapa 3 use el consolidado y los puntajes de la sesión sin volver a subir nada.
"""

import streamlit as st

RAIZ = "hf"


def _vacio():
    return {
        "proceso": {"nombre": "PROC_001_Colombia_Riego", "ronda": 1},
        "etapa1": {
            "df_wide": None,
            "df_alertas": None,
            "resultados": None,
            "excel_bytes": None,
            "nombre_excel": None,
            "timestamp": None,
        },
        "etapa2": {
            "providers": None,
            "items": None,
            "escala": None,
            "puntajes": None,      # dict {"ID|Proveedor": float}
            "timestamp": None,
        },
        "etapa3": {
            "precios_historicos": {},
            "historial": None,
            "tecnica": None,
            "economica": None,
            "workbook_bytes": None,
            "timestamp": None,
        },
    }


def init():
    if RAIZ not in st.session_state:
        st.session_state[RAIZ] = _vacio()
    return st.session_state[RAIZ]


def estado():
    return init()


def reiniciar():
    st.session_state[RAIZ] = _vacio()


# ─── Proceso ──────────────────────────────────────────────────────────────────

def proceso():
    return estado()["proceso"]


def set_proceso(nombre, ronda):
    p = proceso()
    p["nombre"] = nombre
    p["ronda"] = int(ronda)


# ─── Etapas ───────────────────────────────────────────────────────────────────

def etapa(n):
    return estado()[f"etapa{n}"]


def guardar_etapa(n, **campos):
    from datetime import datetime
    e = etapa(n)
    e.update(campos)
    e["timestamp"] = datetime.now()


def etapa1_lista():
    e = etapa(1)
    return e["df_wide"] is not None and not e["df_wide"].empty


def etapa2_lista():
    return bool(etapa(2)["puntajes"])


def etapa3_lista():
    return etapa(3)["workbook_bytes"] is not None
