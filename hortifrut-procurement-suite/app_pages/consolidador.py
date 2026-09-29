"""
Etapa 1 — Consolidador de cotizaciones.
Adaptado de `consolidador_economico/app.py`. Cambios respecto al original:
  - sin st.set_page_config (vive en streamlit_app.py)
  - el Word de feedback se genera con python-docx (sin Node.js)
  - el resultado queda en session_state para que la etapa 3 lo use
"""

import io
import os
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from core.consolidador.core import (
    consolidar_proveedores,
    exportar_consolidado,
    generar_word_feedback,
    parse_supplier_file,
)
from core.shared import estado as E
from core.shared import ui

ui.header(
    "📊 Etapa 1 · Consolidador de cotizaciones",
    "Hortifrut · Procurement · comparación de ofertas y feedback económico",
)
nombre_proceso, ronda = ui.barra_proceso()
ui.flujo(
    "<b>Entra:</b> un Excel por proveedor. &nbsp;·&nbsp; "
    "<b>Sale:</b> el consolidado comparativo, el Word de feedback por proveedor y "
    "los datos que alimentan la etapa 3. &nbsp;·&nbsp; <b>Depende de:</b> nada."
)

with st.sidebar:
    st.markdown("### 📋 Cómo usar")
    st.markdown(
        """
1. Revisa el proceso y la ronda arriba
2. Sube los Excel de cada proveedor
3. Haz clic en **Consolidar**
4. Descarga el Excel y los Word

**Formatos aceptados:** `.xlsx`, `.xls`

**Se detecta automáticamente:**
- Razón social (de Condiciones Comerciales)
- Hoja de cotización, aunque cambie el nombre
- Ítems cotizados vs. no cotizados
- Errores y campos incompletos
"""
    )

# ─── Paso 1 — carga ───────────────────────────────────────────────────────────

col_upload, col_info = st.columns([3, 1])

with col_upload:
    ui.paso("📁 Paso 1 — Sube los archivos de los proveedores")
    archivos = st.file_uploader(
        "Arrastra aquí todos los Excel de proveedores (uno por proveedor)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

with col_info:
    ui.paso("&nbsp;")
    if archivos:
        ui.tarjeta_metrica(len(archivos), "archivo(s) cargado(s)")

st.markdown("<br>", unsafe_allow_html=True)

ya_consolidado = E.etapa1_lista()

if not archivos:
    if ya_consolidado:
        st.success(
            f"Ya hay un consolidado en esta sesión: "
            f"{len(E.etapa(1)['df_wide'])} ítems. Puedes seguir a la etapa 3, "
            "o subir nuevos archivos para reemplazarlo.",
            icon="✅",
        )
        ui.link_pagina("app_pages/escenarios.py", "Ir a la etapa 3 →")
    else:
        st.info("⬆️ Sube al menos un archivo de proveedor para comenzar.", icon="💡")
    st.stop()

col_btn, _ = st.columns([2, 5])
with col_btn:
    consolidar = st.button("🔄 Consolidar proveedores", use_container_width=True,
                            type="primary")

if not consolidar and not ya_consolidado:
    st.stop()

# ─── Paso 2 — procesamiento ───────────────────────────────────────────────────

if consolidar:
    st.markdown("---")
    ui.paso("📥 Paso 2 — Procesando archivos...")

    resultados = []
    progress = st.progress(0)
    tmp_paths = []

    for i, archivo in enumerate(archivos):
        suffix = Path(archivo.name).suffix
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(archivo.getvalue())
        tmp.close()
        tmp_paths.append(tmp.name)

        resultado = parse_supplier_file(
            tmp.name, nombre_proceso=nombre_proceso, ronda=ronda,
            nombre_original=archivo.name,
        )
        resultado["nombre_archivo"] = archivo.name
        resultados.append(resultado)
        progress.progress((i + 1) / len(archivos))

    for p in tmp_paths:
        try:
            os.unlink(p)
        except Exception:
            pass
    progress.empty()

    df_wide, df_alertas = consolidar_proveedores(resultados)

    if df_wide.empty:
        st.error(
            "No se pudo generar la tabla consolidada. Verifica que al menos un "
            "archivo tenga ítems cotizados.", icon="🚨",
        )
        st.stop()

    excel_bytes = exportar_consolidado(df_wide, df_alertas, nombre_proceso, ronda)
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    E.guardar_etapa(
        1,
        df_wide=df_wide,
        df_alertas=df_alertas,
        resultados=resultados,
        excel_bytes=excel_bytes,
        nombre_excel=f"CONSOLIDADO_{nombre_proceso}_R{ronda}_{fecha}.xlsx",
    )

# Datos de la sesión (recién calculados o de una corrida anterior)
e1 = E.etapa(1)
df_wide = e1["df_wide"]
resultados = e1["resultados"] or []
excel_bytes = e1["excel_bytes"]
nombre_salida = e1["nombre_excel"]

# ─── Resumen por proveedor ────────────────────────────────────────────────────

ui.paso("📋 Resumen de archivos procesados")
cols_cards = st.columns(min(max(len(resultados), 1), 4))

for i, r in enumerate(resultados):
    tiene_error = bool(r["errores"])
    tiene_warning = bool(r["warnings"])
    clase = "error" if tiene_error else ("warning" if tiene_warning else "")
    icono = "❌" if tiene_error else ("⚠️" if tiene_warning else "✅")

    with cols_cards[i % len(cols_cards)]:
        st.markdown(
            f"""
        <div class="prov-card {clase}">
            <strong>{icono} {r['proveedor']}</strong><br>
            <small>{r.get('nombre_archivo', '')}</small><br>
            <small><b>{len(r['df'])}</b> ítems cotizados</small>
        </div>
        """,
            unsafe_allow_html=True,
        )
        if tiene_error:
            for e in r["errores"]:
                st.error(f"**Error:** {e}", icon="🚨")
        if tiene_warning:
            with st.expander(f"Ver {len(r['warnings'])} advertencia(s)"):
                for w in r["warnings"]:
                    st.warning(w, icon="⚠️")

# ─── Métricas ─────────────────────────────────────────────────────────────────

st.markdown("---")
ui.paso("📊 Paso 3 — Resultados del benchmarking")

total_items = len(df_wide)
provs_ok = sum(1 for r in resultados if not r["errores"] and len(r["df"]) > 0)

cols_pu = [c for c in df_wide.columns if c.startswith("PU_Total__")]
items_con_todas = int(df_wide[cols_pu].notna().all(axis=1).sum()) if cols_pu else 0

ahorro_potencial = None
if {"PRECIO_MIN", "PRECIO_MAX", "Cantidad_Total"}.issubset(df_wide.columns):
    df_tmp = df_wide.copy()
    df_tmp["Cantidad_Total"] = pd.to_numeric(df_tmp["Cantidad_Total"], errors="coerce").fillna(0)
    ahorro_potencial = ((df_tmp["PRECIO_MAX"] - df_tmp["PRECIO_MIN"]) * df_tmp["Cantidad_Total"]).sum()

m1, m2, m3, m4 = st.columns(4)
with m1:
    ui.tarjeta_metrica(total_items, "Ítems en tabla")
with m2:
    ui.tarjeta_metrica(provs_ok, "Proveedores válidos")
with m3:
    ui.tarjeta_metrica(items_con_todas, "Ítems con todas las ofertas")
with m4:
    val = f"${ahorro_potencial:,.0f}" if ahorro_potencial is not None else "—"
    ui.tarjeta_metrica(val, "Diferencial máx. (USD)", size="1.4rem")

st.markdown("<br>", unsafe_allow_html=True)

with st.expander("📊 Tabla consolidada", expanded=True):
    def color_delta(val):
        if pd.isna(val):
            return ""
        if val <= 5:
            return "background-color: #C6EFCE; color: #375623"
        if val <= 20:
            return "background-color: #FFEB9C; color: #7D6608"
        return "background-color: #FFC7CE; color: #9C0006"

    try:
        styled = df_wide.style
        if "DELTA_%" in df_wide.columns:
            styled = styled.map(color_delta, subset=["DELTA_%"])
        st.dataframe(styled, use_container_width=True, height=400)
    except Exception:
        st.dataframe(df_wide, use_container_width=True, height=400)

    st.caption(
        f"🟢 Delta ≤5% · 🟡 5–20% · 🔴 >20%  |  {total_items} ítems · {provs_ok} proveedores"
    )

# ─── Descarga del Excel ───────────────────────────────────────────────────────

st.markdown("---")
ui.paso("💾 Paso 4 — Descarga el Excel consolidado")

col_dl, col_desc = st.columns([2, 4])
with col_dl:
    st.download_button(
        label="⬇️ Descargar Excel consolidado",
        data=excel_bytes,
        file_name=nombre_salida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_desc:
    st.markdown(
        f"""
    **El archivo incluye:**
    - 📊 Hoja **Consolidado** — tabla con todos los proveedores lado a lado
    - Columnas de benchmarking: precio mín/máx, delta %, proveedor ganador
    - Semaforización automática por precio

    *Nombre del archivo: `{nombre_salida}`*

    Para la etapa 3 **no necesitas este archivo**: los datos ya están en la sesión.
    """
    )

# ─── Word de feedback ─────────────────────────────────────────────────────────

st.markdown("---")
ui.paso("📄 Paso 5 — Feedback económico (Word) por proveedor")
st.caption(
    "Un documento por proveedor con la tabla de tarifas, la semaforización y la "
    "leyenda de dispersión — listo para enviar. Ahora se genera 100% en Python."
)

proveedores_validos = [
    r["proveedor"] for r in resultados if not r["errores"] and len(r["df"]) > 0
]

if not proveedores_validos:
    st.info("No hay proveedores válidos para generar el Word de feedback.")
else:
    docs_generados = {}
    errores_word = {}

    for prov in proveedores_validos:
        word_bytes, error_msg = generar_word_feedback(df_wide, prov, nombre_proceso, ronda)
        if word_bytes:
            nombre_word = f"Feedback_{prov.replace(' ', '_')}_{nombre_proceso}_R{ronda}.docx"
            docs_generados[prov] = (word_bytes, nombre_word)
        else:
            errores_word[prov] = error_msg

    if docs_generados:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for prov, (word_bytes, nombre_word) in docs_generados.items():
                zf.writestr(nombre_word, word_bytes)
        zip_buffer.seek(0)

        st.download_button(
            label=f"📦 Descargar todos los Word ({len(docs_generados)} proveedores)",
            data=zip_buffer.getvalue(),
            file_name=f"Feedback_TODOS_{nombre_proceso}_R{ronda}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )
        st.markdown("<br>", unsafe_allow_html=True)

    with st.expander("O descarga uno por uno"):
        cols_word = st.columns(min(len(proveedores_validos), 4))
        for i, prov in enumerate(proveedores_validos):
            with cols_word[i % len(cols_word)]:
                if prov in docs_generados:
                    word_bytes, nombre_word = docs_generados[prov]
                    st.download_button(
                        label=f"⬇️ {prov}",
                        data=word_bytes,
                        file_name=nombre_word,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key=f"word_{i}",
                    )
                else:
                    st.warning(f"No se pudo generar el Word para {prov}.", icon="⚠️")
                    with st.expander("Ver detalle del error"):
                        st.code(errores_word.get(prov, ""))

st.markdown("---")
st.success("Etapa 1 lista. Los datos quedaron disponibles para la etapa 3.", icon="✅")
col_a, col_b = st.columns(2)
with col_a:
    ui.link_pagina("app_pages/evaluacion_tecnica.py", "Continuar a la etapa 2 (técnica) →")
with col_b:
    ui.link_pagina("app_pages/escenarios.py", "Saltar a la etapa 3 (adjudicación) →")
