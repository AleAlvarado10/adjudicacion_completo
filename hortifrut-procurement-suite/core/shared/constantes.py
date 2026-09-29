"""
Fuente única de verdad para umbrales, semáforos y colores de la suite.

IMPORTANTE: estos valores son EXACTAMENTE los que ya usaban las tres apps
originales. No se cambió ningún criterio de negocio al unificar.
"""

# ─── Semáforo económico (consolidador: Excel y Word de feedback) ──────────────
# Base de cálculo: P.U. EXW del proveedor vs. el mínimo EXW del ítem.
DISPERSION_AZUL = 5       # <= 5%
DISPERSION_VERDE = 25     # <= 25%
DISPERSION_AMARILLO = 50  # <= 50%   (> 50% => Rojo)

# ─── Semáforo de dispersión de la vista DELTA_% (tabla en pantalla) ───────────
DELTA_VERDE = 5
DELTA_AMARILLO = 20

# ─── Semáforo técnico (evaluación técnica y escenarios) ───────────────────────
UMBRAL_TECNICO_ALTO = 3.75   # >= 3.75  ALTO
UMBRAL_TECNICO_MEDIO = 2.5   # >= 2.50  MEDIO   (< 2.50 BAJO)

# ─── Escenarios ───────────────────────────────────────────────────────────────
INDICE_INFLACION_DEFAULT = 0.0614

# ─── Paleta institucional Hortifrut (manual de marca 2025) ────────────────────
AZUL = "#192680"          # Blueberry Blue
AZUL_CLARO = "#3a52b5"    # Blueberry Blue Bright 1
VERDE = "#0c8d50"         # Nature Green
VERDE_CLARO = "#87c845"
MARFIL = "#fcf8de"        # Human Ivory
GRIS_TEXTO = "#3a3a3a"

# Rellenos de semáforo para Excel (openpyxl, sin '#')
FILL_VERDE = "C6EFCE"
FILL_AMARILLO = "FFEB9C"
FILL_ROJO = "FFC7CE"
FILL_AZUL = "BDD7EE"
