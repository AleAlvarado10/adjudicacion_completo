# Hortifrut · Procurement Suite

Las tres herramientas de procurement que antes vivían en tres links de Streamlit
separados, unificadas en **una sola app con un solo link**.

| Etapa | Qué hace | Entra | Sale |
|---|---|---|---|
| **1. Consolidar cotizaciones** | Compara las ofertas de todos los proveedores | Un Excel por proveedor | Excel consolidado + Word de feedback económico por proveedor |
| **2. Evaluación técnica** | Califica técnicamente cada ítem | Matriz de evaluación (.xlsx) | Reporte Word por proveedor + puntajes /5.0 |
| **3. Propuesta de adjudicación** | Escenarios, ahorro y plan de homologación | El consolidado + los puntajes | Workbook de 11 hojas cruzadas por fórmula |
| **Historial de rondas** | Compara precios entre rondas | Historial de la sesión o archivo previo | Comparación ronda a ronda |

```
Excel de proveedores (N)
        │
        ▼
 [1] CONSOLIDADOR ─────► consolidado ────┐
                                         │
 Matriz de evaluación                    │
        │                                ▼
 [2] EVAL. TÉCNICA ────► puntajes ──► [3] ESCENARIOS ──► Prop Adj Final (7 hojas)
```

**Lo importante:** pasar de una etapa a otra ya **no requiere descargar y volver a
subir archivos**. Los datos quedan en la sesión y la etapa siguiente los ofrece
como opción por defecto. El archivo intermedio sigue disponible por si quieres
auditarlo o retomar el trabajo otro día.

---

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Se abre en `http://localhost:8501`.

---

## Desplegar en Streamlit Community Cloud

1. Sube este repo a GitHub (revisa que `.gitignore` esté incluido: **nunca subas
   Excel con precios reales de proveedores**).
2. En [share.streamlit.io](https://share.streamlit.io) → **New app**, elige el repo,
   la rama y como *Main file path*: `streamlit_app.py`.
   En **Advanced settings** elige **Python 3.12**. Si dejas la versión por defecto,
   el despliegue puede quedarse colgado: las versiones más nuevas de Python todavía
   no tienen paquetes precompilados de pandas y el instalador intenta compilarlo.
   El repo incluye `.python-version` con 3.12 por si el instalador lo respeta.
3. **Advanced settings → Secrets**, si quieres contraseña de equipo:
   ```toml
   app_password = "la-clave-del-equipo"
   ```
   Si no defines nada, la app abre sin contraseña.
4. En **Settings → Sharing**, restringe el acceso a los correos del equipo.

Una vez desplegada, **los tres links anteriores quedan obsoletos** y se pueden dar
de baja.

---

## Estructura

```
streamlit_app.py            # único punto de entrada: page_config, navegación, acceso
app_pages/                  # una página por etapa
  inicio.py                 #   mapa del proceso y estado del pipeline
  consolidador.py           #   etapa 1
  evaluacion_tecnica.py     #   etapa 2
  escenarios.py             #   etapa 3
  historial.py              #   historial de rondas
core/
  shared/                   # estado de sesión, constantes y UI comunes
  consolidador/             # parseo, consolidación, Excel y Word económico
  tecnica/                  # parser de la matriz y reporte Word técnico
  escenarios/               # 9 módulos del motor de escenarios (sin cambios de lógica)
.streamlit/config.toml      # tema Hortifrut y límite de subida
```

---

## Qué cambió respecto a las tres apps originales

- **Se eliminó la dependencia de Node.js.** El Word de feedback económico se
  generaba con un script `generar_feedback_word.js` y la librería `docx` de npm;
  Streamlit Cloud no tiene Node, así que se portó a `python-docx`
  (`core/consolidador/word_feedback.py`) manteniendo exactamente las mismas
  medidas, colores y textos. Ya no hacen falta `package.json` ni `node_modules/`.
- **Un solo `st.set_page_config`**, en el entrypoint.
- **Sin estado global mutable.** La escala de puntajes de la evaluación técnica
  vivía en un `global`; ahora se devuelve del parser y se pasa como argumento
  (en Streamlit Cloud varios usuarios comparten el mismo proceso).
- **Proceso y ronda se piden una sola vez** y se propagan a las tres etapas.
- **El historial ya no se escribe junto al código**, sino en `/tmp`, y la UI avisa
  que no es persistente.
- **Factor logístico opcional:** si subes la tabla, todo funciona igual que antes. Si
  no la subes, el archivo se genera **sin** las columnas `INCOTERM`, `Factor Logistico`
  ni `Precio Histórico 2526 + factor`: el Costo total queda igual al Precio Unitario y
  el Baseline se calcula directo sobre el Precio Histórico.
- **Escenario 2 = Mejor Puntaje Técnico.** Antes asignaba el 100% al más barato *solo
  entre los homologados*, así que un ítem sin ningún proveedor homologado quedaba sin
  asignar. Ahora asigna el 100% al proveedor con el **mayor puntaje técnico del ítem**,
  sin importar si ese puntaje es alto, medio o bajo. Si dos empatan en puntaje, gana el
  más barato de los dos.
- **Un solo archivo previo por ronda.** Antes había cinco campos distintos (historial,
  detalle técnico previo, evaluación económica previa, precio histórico y plan). Ahora se
  elige si es la ronda 1 o posterior, y en el segundo caso se sube **solo** el workbook
  de la ronda anterior: de ahí salen los precios de las rondas previas, los puntajes
  técnicos, el precio histórico y los estados del plan.
- **Hojas Estrategia Finalistas y Decisión por Ítem.** La primera muestra cobertura,
  puntaje promedio y monto actual de cada proveedor, con una columna SÍ/NO para marcar
  finalistas: al marcarlos se recalcula cuántos ítems quedan cubiertos, el monto y el
  ahorro de consolidar en ese set. La segunda lista, ítem por ítem, quién gana en cada
  escenario y con qué monto.
- **Orden de hojas** igual al workbook de referencia: Léeme, Prop Adj Final, Resumen,
  Resumen Ejecutivo, Estrategia Finalistas, Decisión por Ítem, y luego el detalle.
- **Hoja Resumen Ejecutivo nueva.** Monto adjudicado por destino con el proveedor
  principal, qué proveedor queda en cada destino con subtotales, y la lista de ítems a
  decidir en la reunión: aquellos donde el más económico y el de mejor puntaje técnico
  no coinciden, con una columna para marcar la decisión.
- **Hoja Resumen nueva.** El workbook abre con un resumen ejecutivo: comparación de los
  tres escenarios (ítems, proveedores, monto, ahorro y % de ahorro), reparto por
  proveedor y estado técnico. Todo con fórmulas, así que se actualiza si editas el
  Escenario 3 o corriges un puntaje.
- **Precio histórico dentro de la app.** Tabla editable por ítem en la etapa 3, más un
  uploader opcional que detecta solo la columna del ID y la del precio. Antes solo se
  podía escribir en el Excel descargado.
- **Archivos subidos a prueba de recargas.** Streamlit reejecuta el script en cada
  interacción y el cursor del `UploadedFile` queda al final: la segunda lectura devolvía
  vacío y openpyxl fallaba con `BadZipFile`. Ahora todo pasa por `core/shared/archivos.py`.
- **Inflación por país automática.** Se sube una tabla CPI (país | tasa) y cada ítem
  recibe el índice de su país de destino. El índice único manual queda como respaldo.
- **Diagnóstico de cruce de puntajes:** cuando la etapa 3 usa los puntajes de la
  etapa 2, informa cuántas claves `ID|Proveedor` cruzaron y muestra las que no.
  Antes un cruce fallido no se distinguía de "todavía no evaluado".

**Ninguna regla de negocio cambió:** umbrales, fórmulas de openpyxl, semáforos y
formatos de salida son los mismos que en las apps originales.

---

## Limitaciones conocidas (heredadas del código original)

- El historial de rondas **no es persistente** en Streamlit Cloud. Descárgalo al
  cerrar cada ronda y vuelve a subirlo en la siguiente.
- La etapa 2 solo reconoce ítems con ID del tipo `RIEG###`. Para otros procesos
  hay que generalizar el patrón.
- Los nombres de proveedor deben coincidir entre la matriz técnica y las
  cotizaciones para que los puntajes crucen. La etapa 3 te avisa cuando no pasa.
- La hoja **Alertas** del Excel consolidado está desactivada en el código; las
  advertencias se muestran en pantalla durante la etapa 1.
- Si dos proveedores comparten los primeros 20 caracteres de su razón social, el
  Excel consolidado puede fallar al armar la tabla.
