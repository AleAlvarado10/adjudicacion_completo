"""
Etapa 3 — Propuesta de adjudicación y escenarios.
Adaptado de `Evaluacion de escenarios v2/app.py`. Cambios respecto al original:
  - sin st.set_page_config
  - el Consolidado puede venir de la sesión (etapa 1) o subirse a mano
  - los puntajes técnicos pueden venir de la sesión (etapa 2) o de un archivo
  - el historial se guarda en /tmp (Streamlit Cloud no tiene disco persistente)
La lógica de cálculo, las fórmulas y el formato del workbook no se tocaron.
"""

import io
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

from core.escenarios.ahorro import (
    enriquecer, quitar_columnas_auditoria, quitar_columnas_logistica,
)
from core.escenarios.estilo import (
    DATA_START_ROW, HEADER_ROW, aplicar_formato_condicional,
    escribir_encabezado_y_estilo, formatear_datos,
)
from core.escenarios.evaluacion import (
    UMBRAL_ALTO_DEFAULT, UMBRAL_MEDIO_DEFAULT, construir_detalle_evaluacion_tecnica,
    construir_evaluacion_economica, construir_plan_homologacion,
)
from core.escenarios.formulas import escribir_formulas_motor
from core.escenarios.inflacion import load_inflacion
from core.escenarios.precio_historico import leer_precio_historico_auto
from core.escenarios.ronda_previa import leer_ronda_previa, puntajes_desde_tecnica
from core.escenarios.hojas_extra import (
    escribir_consolidado_echo, escribir_detalle_evaluacion_tecnica,
    escribir_detalle_ofertas, escribir_evaluacion_economica, escribir_leeme,
    escribir_plan_homologacion, escribir_resumen,
)
from core.escenarios.logistica import load_factor_logistico, paises_disponibles
from core.escenarios.estrategia import (
    escribir_decision_por_item, escribir_estrategia_finalistas,
)
from core.escenarios.resumen_ejecutivo import escribir_resumen_ejecutivo
from core.escenarios.rondas import guardar_historial_local, historial_wide
from core.escenarios.transformer import (
    consolidado_to_prop_adj_final, load_consolidado, melt_consolidado_completo,
)
from core.shared.archivos import buffer, nombre_de
from core.shared import constantes as C
from core.shared import estado as E
from core.shared import ui

HISTORIAL_PATH_TMP = "/tmp/historial_rondas.xlsx"

ui.header(
    "📈 Etapa 3 · Propuesta de adjudicación + Plan de homologación",
    "Léeme · Consolidado · Detalle Ofertas · Evaluación Técnica · Evaluación Económica · "
    "Plan Homologación · Prop Adj Final",
)
nombre_proceso, ronda_proceso = ui.barra_proceso()
ui.flujo(
    "<b>Entra:</b> el consolidado (de la etapa 1 o subido) más los puntajes técnicos "
    "(de la etapa 2). &nbsp;·&nbsp; <b>Sale:</b> el workbook de 7 hojas "
    "cruzadas por fórmula y el historial de rondas. &nbsp;·&nbsp; "
    "<b>Depende de:</b> etapa 1 (obligatoria) y etapa 2 (opcional)."
)


for key in ["historial_df", "historial_tecnico_df", "historial_economico_df"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ─── Sidebar: parámetros ──────────────────────────────────────────────────────

with st.sidebar:
    st.header("2️⃣ Factor logístico")
    st.caption("Opcional. Si no lo subes, el archivo se genera sin las columnas de factor.")
    factor_file = st.file_uploader("Tabla de Factor Logístico (.xlsx)", type=["xlsx"],
                                    key="factor_file")
    factor_sheet = st.text_input("Hoja de la tabla de factor logístico", value="Hoja1")
    destino_default = st.text_input(
        "País o ciudad de destino por defecto",
        help="Se usa cuando el ítem no trae 'ZONA / LOTE' con el destino. "
             "Puede ser un país (ej. PERU) o una ciudad (ej. TRUJILLO).",
    )

    st.header("3️⃣ Inflación por país")
    st.caption("Tabla CPI: país en la primera columna, tasa en la segunda (ej. PERU | 2,2%).")
    inflacion_file = st.file_uploader("Tabla de inflación (.xlsx)", type=["xlsx"],
                                       key="inflacion_file")
    indice_manual = st.number_input(
        "Índice único (se usa solo si no subes la tabla)",
        value=C.INDICE_INFLACION_DEFAULT, step=0.0001, format="%.4f",
        help="Baseline = Precio Histórico × (1 + Factor Logístico) × (1 + Índice de "
             "inflación). Con la tabla, el índice se asigna solo según el País destino.",
    )

    st.header("4️⃣ Umbrales técnicos")
    umbral_alto = st.number_input("Umbral ALTO (homologa) ≥", value=UMBRAL_ALTO_DEFAULT,
                                   step=0.05, format="%.2f")
    umbral_medio = st.number_input("Umbral MEDIO (por revisar) ≥", value=UMBRAL_MEDIO_DEFAULT,
                                    step=0.05, format="%.2f")

    st.header("5️⃣ ¿Qué ronda es esta?")
    es_primera = st.radio(
        "Ronda del proceso",
        ["Es la primera ronda (R1)", "Es la ronda 2 o posterior"],
        label_visibility="collapsed",
    ) == "Es la primera ronda (R1)"

    ronda_previa_file = None
    ronda_manual = 1
    if es_primera:
        st.caption("No hay nada previo que cargar: la app parte de cero.")
    else:
        ronda_manual = st.number_input("Número de esta ronda", min_value=2, max_value=99,
                                        value=2, step=1)
        ronda_previa_file = st.file_uploader(
            "Archivo de la ronda anterior (PropAdjFinal_....xlsx)", type=["xlsx"],
            key="ronda_previa",
        )
        st.caption(
            "Sube el workbook que descargaste la ronda pasada. De ahí se sacan solos: "
            "los precios de las rondas anteriores, los puntajes técnicos ya evaluados, "
            "el precio histórico y el plan de homologación."
        )

# ─── Origen del Consolidado ───────────────────────────────────────────────────

ui.paso("1️⃣ Consolidado de origen")

tiene_sesion = E.etapa1_lista()
opciones = []
if tiene_sesion:
    opciones.append("Usar el consolidado de esta sesión (etapa 1)")
opciones.append("Subir un Excel Consolidado")

origen = st.radio("¿De dónde tomamos el consolidado?", opciones,
                   label_visibility="collapsed")

df_consolidado = None
nombre_origen = ""

if origen.startswith("Usar el consolidado"):
    e1 = E.etapa(1)
    df_consolidado = e1["df_wide"]
    nombre_origen = e1["nombre_excel"] or "consolidado_sesion.xlsx"
    st.success(
        f"Consolidado de la etapa 1: **{len(df_consolidado)} ítems** · archivo de "
        f"referencia `{nombre_origen}`", icon="✅",
    )
else:
    sheet_name = st.text_input("Nombre de la hoja a leer (vacío = primera hoja)", value="")
    uploaded_file = st.file_uploader("Archivo Excel Consolidado (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        try:
            df_consolidado = load_consolidado(buffer(uploaded_file), sheet_name=sheet_name or None)
            nombre_origen = uploaded_file.name
        except Exception as e:
            st.error(f"⚠️ No se pudo leer el Consolidado: {e}")

if df_consolidado is None:
    if not tiene_sesion:
        st.info(
            "Sube un Excel Consolidado, o corre primero la etapa 1 para que los datos "
            "lleguen automáticamente."
        )
        ui.link_pagina("app_pages/consolidador.py", "Ir a la etapa 1 →")
    st.stop()

ronda_actual = 1 if es_primera else int(ronda_manual)
E.set_proceso(nombre_proceso, ronda_actual)
st.caption(
    f"Procesando como **ronda {ronda_actual}**"
    + ("" if es_primera else " · los precios de las rondas anteriores salen del archivo previo")
)

# ─── Procesar el consolidado ──────────────────────────────────────────────────

try:
    with st.spinner("Procesando Consolidado..."):
        largo, providers = melt_consolidado_completo(df_consolidado)
        rondas = list(range(1, int(ronda_actual) + 1))
        prop_adj, providers_pa, extra_fields = consolidado_to_prop_adj_final(
            df_consolidado, rondas=rondas)
except Exception as e:
    st.error(f"⚠️ No se pudo procesar el Consolidado: {e}")
    st.stop()

st.success(
    f"Base generada: {len(prop_adj)} filas (ítem × proveedor) · "
    f"{len(providers)} proveedores detectados."
)
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Proveedores detectados:**")
    st.write(", ".join(providers))
with col2:
    st.markdown("**Campos extra detectados:**")
    st.write(", ".join(extra_fields) if extra_fields else "ninguno")

# ─── Ronda anterior: un solo archivo del que sale todo ────────────────────────

previa = {"historial": None, "tecnica": None, "economica": None,
          "precios_historicos": {}, "plan_estados": None}

if not es_primera:
    if ronda_previa_file is None:
        st.warning(
            "Indicaste que esta es la ronda "
            f"{int(ronda_manual)}, pero falta el archivo de la ronda anterior. "
            "Súbelo en el panel lateral para no perder los precios, los puntajes y el "
            "precio histórico ya cargados. Si lo dejas vacío, la app sigue igual pero "
            "esta ronda parte sin historial.",
            icon="⚠️",
        )
    else:
        try:
            previa, res_previa = leer_ronda_previa(
                buffer(ronda_previa_file), nombre_archivo=nombre_de(ronda_previa_file))
            detalles = []
            if res_previa["registros_historial"]:
                rr = ", ".join("R" + str(n) for n in res_previa["rondas"])
                detalles.append(f"{res_previa['registros_historial']} precios ({rr})")
            if res_previa["puntajes"]:
                detalles.append(f"{res_previa['puntajes']} puntajes técnicos")
            if res_previa["precios_historicos"]:
                detalles.append(f"{res_previa['precios_historicos']} precios históricos")
            if res_previa["estados_plan"]:
                detalles.append(f"{res_previa['estados_plan']} estados del plan")
            if detalles:
                st.success(
                    "Ronda anterior cargada: " + " · ".join(detalles) +
                    ". No tienes que subir nada más.", icon="✅",
                )
            else:
                st.warning(
                    "El archivo se abrió pero no se encontraron datos aprovechables. "
                    "Revisa que sea el workbook «PropAdjFinal» que genera esta app.",
                    icon="⚠️",
                )
            if res_previa["hojas_faltantes"]:
                st.caption("Hojas no encontradas: " + ", ".join(res_previa["hojas_faltantes"]))
        except Exception as e:
            st.error(f"No se pudo leer el archivo de la ronda anterior: {e}")

# ─── Puntajes técnicos: sesión o archivo ──────────────────────────────────────

st.markdown("### 🔗 Puntajes técnicos")

puntajes_externos = None
puntajes_sesion = E.etapa(2)["puntajes"]

if puntajes_sesion:
    usar_sesion = st.checkbox(
        f"Usar los {len(puntajes_sesion)} puntajes de la etapa 2 de esta sesión",
        value=True,
    )
    if usar_sesion:
        puntajes_externos = dict(puntajes_sesion)

        # Diagnóstico del cruce: la llave es "ID|Proveedor" en ambos lados.
        claves_base = set(largo["Clave"]) if "Clave" in largo.columns else set()
        cruzadas = [k for k in puntajes_externos if k in claves_base]
        no_cruzadas = [k for k in puntajes_externos if k not in claves_base]

        if not claves_base:
            st.warning("No se pudieron construir las claves del consolidado.", icon="⚠️")
        elif not cruzadas:
            st.error(
                f"**Ninguno de los {len(puntajes_externos)} puntajes cruzó con el "
                "consolidado.** Suele ser porque los IDs o los nombres de proveedor "
                "se escriben distinto en la matriz técnica y en las cotizaciones. "
                "Revisa la tabla de abajo y corrige la matriz, o sube los puntajes "
                "como archivo usando el selector del panel lateral.",
                icon="🚨",
            )
        elif no_cruzadas:
            st.warning(
                f"Cruzaron **{len(cruzadas)} de {len(puntajes_externos)}** puntajes. "
                f"{len(no_cruzadas)} no encontraron su par ID|Proveedor.", icon="⚠️",
            )
        else:
            st.success(
                f"Cruzaron los {len(cruzadas)} puntajes de la etapa 2.", icon="✅")

        if no_cruzadas:
            with st.expander(f"Ver las {len(no_cruzadas)} claves que no cruzaron"):
                st.caption(
                    "Izquierda: lo que trae la matriz técnica. Derecha: ejemplos de "
                    "claves válidas del consolidado."
                )
                cc1, cc2 = st.columns(2)
                with cc1:
                    st.dataframe(
                        pd.DataFrame({"Clave de la matriz técnica": sorted(no_cruzadas)}),
                        use_container_width=True, height=260,
                    )
                with cc2:
                    st.dataframe(
                        pd.DataFrame({"Claves del consolidado": sorted(claves_base)[:200]}),
                        use_container_width=True, height=260,
                    )

if puntajes_externos is None and previa["tecnica"] is not None:
    recuperados = puntajes_desde_tecnica(previa["tecnica"])
    if recuperados:
        puntajes_externos = recuperados
        st.success(
            f"Se recuperaron {len(recuperados)} puntajes técnicos del archivo de la "
            "ronda anterior.", icon="✅",
        )

if not puntajes_externos:
    st.info(
        "No hay puntajes técnicos en esta sesión. Corre la etapa 2, o complétalos a "
        "mano en la hoja «Detalle Evaluación Técnica» del workbook que descargues."
    )

# ─── Evaluación técnica y económica (carry-forward) ───────────────────────────

historial_tecnico = previa["tecnica"] if previa["tecnica"] is not None else st.session_state["historial_tecnico_df"]
historial_economico = previa["economica"] if previa["economica"] is not None else st.session_state["historial_economico_df"]
plan_estados = previa["plan_estados"]

tecnica = construir_detalle_evaluacion_tecnica(
    largo, historial_tecnico=historial_tecnico, plan_estados=plan_estados,
    puntajes_externos=puntajes_externos)

# ─── Precio histórico: se escribe acá mismo, en la tabla ──────────────────────

st.markdown("### 💰 Precio histórico por ítem")
st.caption(
    "Es lo que se pagó la temporada pasada. Es el único dato que la app no puede sacar "
    "de las cotizaciones. Sirve para calcular el Baseline y, con eso, el ahorro."
)

ids_proceso = (
    largo[["ID", "Descripcion"]].drop_duplicates(subset=["ID"])
    if {"ID", "Descripcion"}.issubset(largo.columns)
    else largo[["ID"]].drop_duplicates().assign(Descripcion="")
)
ids_proceso["ID"] = ids_proceso["ID"].astype(str).str.strip()

# Prioridad: lo que el usuario ya escribió en esta sesión > archivo de la ronda
# anterior > archivo externo que suba ahora.
precios_guardados = dict(E.etapa(3).get("precios_historicos") or {})
precios_guardados.update({k: v for k, v in (previa["precios_historicos"] or {}).items()
                          if k not in precios_guardados})

col_carga, col_info = st.columns([2, 3])
with col_carga:
    historico_file = st.file_uploader(
        "¿Lo tienes en un Excel? Súbelo acá (opcional)", type=["xlsx"],
        key="historico_auto",
        help="Detecta solo la columna del ID y la del precio. No tienes que elegir "
             "hoja ni columna.",
    )
if historico_file is not None:
    try:
        leidos, res_h = leer_precio_historico_auto(
            buffer(historico_file), ids_validos=ids_proceso["ID"].tolist())
        precios_guardados.update(leidos)
        with col_info:
            st.success(
                f"Leídos **{res_h['leidos']}** precios de la hoja «{res_h['hoja']}», "
                f"cruzando la columna «{res_h['col_id']}» con «{res_h['col_precio']}»."
                + (f" Se ignoraron {res_h['sin_match']} IDs que no son de este proceso."
                   if res_h["sin_match"] else ""),
                icon="✅",
            )
        faltantes = [i for i in ids_proceso["ID"] if i not in precios_guardados]
        if faltantes:
            st.caption(
                f"{len(faltantes)} ítems del proceso no traen precio en el archivo: "
                "puedes escribirlos a mano en la tabla de abajo."
            )
    except Exception as e:
        with col_info:
            st.error(f"No se pudo leer el archivo: {e}")

tabla_precios = ids_proceso.copy()
tabla_precios["Precio Histórico (USD)"] = [
    float(precios_guardados.get(i, 0) or 0) for i in tabla_precios["ID"]
]
tabla_precios = tabla_precios.rename(columns={"Descripcion": "Descripción"})

editada = st.data_editor(
    tabla_precios,
    use_container_width=True,
    hide_index=True,
    key="editor_historico",
    column_config={
        "ID": st.column_config.TextColumn("ID", disabled=True, width="small"),
        "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
        "Precio Histórico (USD)": st.column_config.NumberColumn(
            "Precio Histórico (USD)", min_value=0.0, step=0.01, format="%.2f",
            help="Escribe el precio de la temporada pasada. Deja 0 si no lo tienes.",
        ),
    },
)

precios_externos = {
    str(r["ID"]).strip(): float(r["Precio Histórico (USD)"])
    for _, r in editada.iterrows()
    if pd.notna(r["Precio Histórico (USD)"]) and float(r["Precio Histórico (USD)"]) > 0
} or None

E.guardar_etapa(3, precios_historicos=precios_externos or {})

n_con = len(precios_externos or {})
if n_con:
    st.caption(
        f"{n_con} de {len(tabla_precios)} ítems con precio histórico. "
        "Los que quedan en 0 usan el costo de la Ronda 1 como referencia."
    )
else:
    st.caption(
        "Sin precio histórico: el Baseline usará el costo de la Ronda 1 como "
        "referencia. También puedes completarlo después en la hoja «Evaluación "
        "Económica» del Excel que descargues."
    )

economica = construir_evaluacion_economica(
    largo, historial_economico=historial_economico, precios_externos=precios_externos)
plan = construir_plan_homologacion(largo, providers)

n_desde_plan = (tecnica["Estado desde Plan"].notna().sum()
                if "Estado desde Plan" in tecnica.columns else 0)
n_puntajes = tecnica["Puntaje Técnico (/5.0)"].notna().sum()
n_historicos = (economica["Precio Histórico (crudo)"] > 0).sum()
st.markdown(
    f"**Evaluación Técnica:** {n_puntajes} de {len(tecnica)} filas ya tienen puntaje. "
    f"**Evaluación Económica:** {n_historicos} de {len(economica)} ítems con Precio "
    "Histórico. "
    + (f"**Plan de Homologación:** {n_desde_plan} filas con estado importado "
       "(manda sobre el puntaje)." if n_desde_plan else "")
)

# ─── Factor logístico + historial + ahorro ────────────────────────────────────

mapa_inflacion = None
if inflacion_file is not None:
    try:
        mapa_inflacion, res_infl = load_inflacion(buffer(inflacion_file))
        st.info(
            f"Tabla de inflación cargada: {res_infl['paises']} países. "
            "El índice se asigna por «País destino» de cada ítem.", icon="🌐",
        )
    except Exception as e:
        st.error(f"⚠️ No se pudo leer la tabla de inflación: {e}")

indice_inflacion = mapa_inflacion if mapa_inflacion else indice_manual

lookup_logistico = None
if factor_file is not None:
    try:
        lookup_logistico = load_factor_logistico(buffer(factor_file), sheet_name=factor_sheet)
        st.info(
            f"Tabla de factor logístico cargada: {len(lookup_logistico)} rutas · "
            f"países disponibles: {', '.join(paises_disponibles(lookup_logistico))}"
        )
    except Exception as e:
        st.error(f"⚠️ No se pudo leer la tabla de factor logístico: {e}")

historial_previo = previa["historial"]
if historial_previo is None and st.session_state["historial_df"] is not None:
    historial_previo = st.session_state["historial_df"]

with st.spinner("Aplicando factor logístico, historial de rondas y ahorro..."):
    enriquecido, historial_actualizado, resumen = enriquecer(
        prop_adj,
        lookup_logistico=lookup_logistico,
        destino_default=destino_default or None,
        historial_df=historial_previo,
        ronda_actual=int(ronda_actual),
        archivo_origen=nombre_origen,
        indice_inflacion=indice_inflacion,
        rondas=rondas,
    )

st.session_state["historial_df"] = historial_actualizado
st.session_state["historial_tecnico_df"] = tecnica
st.session_state["historial_economico_df"] = economica
guardar_historial_local(historial_actualizado, path=HISTORIAL_PATH_TMP)
E.guardar_etapa(3, historial=historial_actualizado, tecnica=tecnica, economica=economica)

if isinstance(indice_inflacion, dict):
    sin_pais = resumen.get("inflacion_sin_pais") or []
    resueltas = resumen.get("inflacion_resueltas", 0)
    if sin_pais:
        st.warning(
            f"Inflación asignada en {resueltas} de {len(enriquecido)} filas. "
            f"Estos destinos no están en la tabla CPI y quedaron en 0: "
            f"**{', '.join(sin_pais)}**. Agrégalos a la tabla o corrige el país en "
            "la cotización.", icon="⚠️",
        )
    else:
        st.success(
            f"Inflación asignada por país en las {resueltas} filas.", icon="✅")

if lookup_logistico is not None:
    st.markdown("### 🌎 Resultado del cruce con Factor Logístico")
    estados = resumen.get("estados_logistico", {})
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Resueltos (OK)", estados.get("OK", 0))
    c2.metric("❓ Requieren destino más específico", estados.get("REQUIERE_DESTINO", 0))
    c3.metric("🚫 Proveedor no encontrado", estados.get("PROVEEDOR_NO_ENCONTRADO", 0))
    if estados.get("PROVEEDOR_NO_ENCONTRADO", 0) or estados.get("REQUIERE_DESTINO", 0):
        no_ok = enriquecido[enriquecido["Estado_Logistico"] != "OK"]
        proveedores_problema = sorted(
            no_ok["Nombre del proveedor"].dropna().unique().tolist())
        st.warning(
            "Proveedores que no quedaron 100% resueltos: "
            f"**{', '.join(proveedores_problema)}**"
        )

st.markdown(
    f"**Rondas representadas:** {', '.join('R' + str(n) for n in resumen.get('rondas', []))} "
    f"· ronda vigente: **R{resumen.get('ronda_vigente')}** (es la que alimenta Mínimo, "
    "Desviación, Escenarios y Ahorro)."
)

wide = historial_wide(historial_actualizado)
if not wide.empty:
    st.markdown("### Historial de rondas (comparación)")
    st.dataframe(wide, use_container_width=True)

con_logistica = lookup_logistico is not None


def hoja_final(df):
    """Hoja Prop Adj Final: sin auditoría y, si no hay tabla logística, sin factor."""
    out = quitar_columnas_auditoria(df)
    if not con_logistica:
        out = quitar_columnas_logistica(out)
    return out


if not con_logistica:
    st.info(
        "No subiste tabla de Factor Logístico, así que el archivo se genera **sin** las "
        "columnas INCOTERM, Factor Logistico y «Precio Histórico 2526 + factor». "
        "El Costo total queda igual al Precio Unitario y el Baseline se calcula directo "
        "sobre el Precio Histórico. Si subes la tabla en el panel lateral, vuelven a "
        "aparecer y todo se comporta como antes.",
        icon="ℹ️",
    )

st.markdown("### Vista previa — Prop Adj Final")
st.dataframe(hoja_final(enriquecido), use_container_width=True)

with st.expander("Vista previa — Detalle Evaluación Técnica"):
    st.dataframe(tecnica, use_container_width=True)
with st.expander("Vista previa — Evaluación Económica"):
    st.dataframe(economica, use_container_width=True)
with st.expander("Vista previa — Plan Homologación"):
    st.dataframe(plan, use_container_width=True)

# ─── Workbook de 7 hojas ──────────────────────────────────────────────────────


def build_workbook(ronda_actual: int) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    factor_modo = "tabla" if lookup_logistico is not None else "sin_tabla"
    escribir_leeme(wb, ronda_actual, umbral_alto, umbral_medio, factor_modo, providers,
                    indice_inflacion=indice_inflacion, rondas=rondas)
    escribir_consolidado_echo(wb, df_consolidado)
    escribir_detalle_ofertas(wb, largo)
    ws_t, t_ini, t_fin = escribir_detalle_evaluacion_tecnica(
        wb, tecnica, umbral_alto, umbral_medio)
    ws_e, e_ini, e_fin = escribir_evaluacion_economica(wb, economica)
    escribir_plan_homologacion(wb, plan, providers, (t_ini, t_fin))

    ws_p = wb.create_sheet("Prop Adj Final")
    hoja_df = hoja_final(enriquecido)
    for c, col_name in enumerate(hoja_df.columns, start=1):
        ws_p.cell(row=HEADER_ROW, column=c, value=col_name)
    for i, (_, row) in enumerate(hoja_df.iterrows()):
        r = DATA_START_ROW + i
        for c, val in enumerate(row, start=1):
            ws_p.cell(row=r, column=c, value=(None if pd.isna(val) else val))

    umbrales = escribir_encabezado_y_estilo(
        ws_p, hoja_df,
        titulo=f"PROPUESTA ADJUDICACIÓN FINAL  |  Ronda {ronda_actual}  |  "
               "generado automáticamente (USD)",
    )
    escribir_formulas_motor(
        ws_p, hoja_df, header_row=HEADER_ROW, data_start_row=DATA_START_ROW,
        tecnica_rango=(t_ini, t_fin), eval_economica_rango=(e_ini, e_fin),
        rondas=rondas,
    )
    # La hoja Resumen va al final para poder apuntar a las filas ya escritas,
    # pero se inserta en la posición 2 del libro (justo después de Léeme).
    rango_prop = (DATA_START_ROW, DATA_START_ROW + len(hoja_df) - 1)
    n_items = hoja_df["ID SAP"].nunique() if "ID SAP" in hoja_df.columns else 0

    _, dec_rango = escribir_decision_por_item(
        wb, hoja_df, prop_rango=rango_prop, columnas_prop=list(hoja_df.columns),
        ronda_vigente=max(rondas), providers=providers,
    )
    if dec_rango:
        escribir_estrategia_finalistas(
            wb, providers, prop_rango=rango_prop,
            columnas_prop=list(hoja_df.columns), decision_rango=dec_rango,
            ronda_vigente=max(rondas), total_items=n_items,
        )

    escribir_resumen_ejecutivo(
        wb, hoja_df,
        prop_rango=(DATA_START_ROW, DATA_START_ROW + len(hoja_df) - 1),
        columnas_prop=list(hoja_df.columns),
        ronda_actual=ronda_actual, ronda_vigente=max(rondas),
        nombre_proceso=nombre_proceso,
    )
    escribir_resumen(
        wb, providers,
        prop_rango=(DATA_START_ROW, DATA_START_ROW + len(hoja_df) - 1),
        columnas_prop=list(hoja_df.columns),
        ronda_actual=ronda_actual, nombre_proceso=nombre_proceso,
    )

    orden = ["Léeme", "Prop Adj Final", "Resumen", "Resumen Ejecutivo",
             "Estrategia Finalistas", "Decisión por Ítem", "Consolidado",
             "Detalle Ofertas", "Detalle Evaluación Técnica", "Evaluación Económica",
             "Plan Homologación"]
    wb._sheets = ([wb[n] for n in orden if n in wb.sheetnames]
                  + [h for h in wb._sheets if h.title not in orden])

    formatear_datos(ws_p, hoja_df)
    aplicar_formato_condicional(
        ws_p, hoja_df,
        verde_cell=umbrales["verde_cell"], amarillo_cell=umbrales["amarillo_cell"])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


output_bytes = build_workbook(int(ronda_actual))
E.guardar_etapa(3, workbook_bytes=output_bytes)
fname = f"PropAdjFinal_Ronda{int(ronda_actual)}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

st.markdown("---")
ui.paso("💾 Descargas")
col_a, col_b = st.columns(2)
with col_a:
    st.download_button(
        label="⬇️ Descargar workbook completo (.xlsx)",
        data=output_bytes, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_b:
    hist_buffer = io.BytesIO()
    with pd.ExcelWriter(hist_buffer, engine="openpyxl") as writer:
        historial_actualizado.to_excel(writer, sheet_name="Historial", index=False)
    hist_buffer.seek(0)
    st.download_button(
        label="⬇️ Descargar historial de precios (.xlsx)",
        data=hist_buffer.getvalue(), file_name="historial_rondas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.caption(
    "El archivo trae 11 hojas: Léeme, Prop Adj Final, Resumen, Resumen Ejecutivo, Estrategia Finalistas, Decisión por Ítem, Consolidado, Detalle Ofertas, Detalle "
    "Evaluación Técnica, Evaluación Económica, Plan Homologación y Prop Adj Final — "
    "todas cruzadas por fórmula. Completa 'Puntaje Técnico' y 'Precio Histórico "
    "(crudo)' y Homologación, Escenario 2, semáforos y Ahorro recalculan solos. "
    "Para la próxima ronda, marca «Es la ronda 2 o posterior» en el panel lateral y "
    "sube este mismo archivo: de ahí salen solos los precios anteriores, los puntajes, "
    "el precio histórico y el plan."
)
