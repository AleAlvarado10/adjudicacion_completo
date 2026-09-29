"""
Fórmulas 'motor' de la hoja Prop Adj Final (estructura nueva).

Todo se escribe como FÓRMULAS de Excel, para que recalculen solas si se
edita cualquier dato de entrada (Factor Logistico, Precio Unitario de
cualquier ronda, Precio Histórico, Índice de inflación, % Adj. E3):

  Costo total (USD) R{n} = Precio Unitario R{n} × (1 + Factor Logistico)
  Mínimo costo ofertado  = MINIFS(Costo total ronda vigente, ">0", por ID SAP)
  Desviación con Mín [%] = (Costo total vigente / Mínimo) − 1
  Calidad / Homologación = cruce por Clave (ID|Proveedor) con 'Detalle Ev. Técnica'
  Puntaje Técnico        = idem
  Precio Histórico 2526  = cruce por ID con 'Evaluación Económica'
  Precio Hist. + factor  = Precio Histórico × (1 + Factor Logistico)
  Baseline               = (Precio Hist. + factor) × (1 + Índice inflación),
                           y si no hay histórico -> Costo total R1

  Escenario 1 (auto)  = 100% al Costo total vigente más bajo del ítem
  Escenario 2 (auto)  = 100% al proveedor con el MAYOR puntaje técnico del ítem,
                        sin importar si ese puntaje es alto, medio o bajo. Si dos
                        empatan en puntaje, gana el más barato de los dos.
  Escenario 3 (manual)= % que el usuario complete (parte en 0)

Si NO se subió tabla de Factor Logístico, las columnas INCOTERM, Factor Logistico
y 'Precio Histórico 2526 + factor' no existen en la hoja: el Costo total pasa a ser
igual al Precio Unitario y el Baseline se calcula directo sobre el Precio Histórico.

  Vol. E{n}    = % Adj. E{n} × Cantidad Esperada
  Monto E{n}   = Vol. E{n} × Costo total vigente
  Ahorro_E{n}  = (Baseline − Costo total vigente) × Vol. E{n}
"""

from openpyxl.utils import get_column_letter

H_ID = "ID SAP"
H_PAIS = "País destino"
H_CANTIDAD = "Cantidad Esperada"
H_PROVEEDOR = "Nombre del proveedor"
H_FACTOR = "Factor Logistico"
H_MINIMO = "Mínimo costo\nofertado"
H_DESVIACION = "Desviación\ncon Mín [%]"
H_HOMOLOGACION = "Calidad /\nHomologación"
H_PUNTAJE = "Puntaje Técnico\n(/5.0)"
H_HISTORICO = "Precio Histórico 2526"
H_HISTORICO_FACTOR = "Precio Histórico 2526 + factor"
H_INFLACION = "Indice de inflación"
H_BASELINE = "Baseline"

HOMOLOGADO = "HOMOLOGADO"


def _col(df, header):
    return df.columns.get_loc(header) + 1 if header in df.columns else None


def _L(df, header):
    c = _col(df, header)
    return get_column_letter(c) if c else None


def col_precio_unitario(n):
    return f"Precio Unitario R{n}\n(USD)"


def col_costo_total(n):
    return f"Costo total  (USD) R{n}"


def escribir_formulas_motor(ws, df, header_row=3, data_start_row=4,
                             tecnica_rango=None, eval_economica_rango=None,
                             rondas=(1,)):
    """Escribe las fórmulas del motor de análisis en la hoja."""
    n_rows = len(df)
    last_row = data_start_row + n_rows - 1
    if n_rows == 0:
        return

    L = lambda h: _L(df, h)  # noqa: E731

    rondas = list(rondas) or [1]
    ronda_vigente = rondas[-1]
    ronda_base = rondas[0]

    id_l = L(H_ID)
    prov_l = L(H_PROVEEDOR)
    factor_l = L(H_FACTOR)
    cantidad_l = L(H_CANTIDAD)
    min_l, desv_l = L(H_MINIMO), L(H_DESVIACION)
    homolog_l, puntaje_l = L(H_HOMOLOGACION), L(H_PUNTAJE)
    hist_l = L(H_HISTORICO)
    hist_factor_l = L(H_HISTORICO_FACTOR)
    infl_l = L(H_INFLACION)
    baseline_l = L(H_BASELINE)

    total_vigente_l = L(col_costo_total(ronda_vigente))
    total_base_l = L(col_costo_total(ronda_base))

    def rng(col_letter):
        return f"${col_letter}${data_start_row}:${col_letter}${last_row}"

    for r in range(data_start_row, last_row + 1):
        # --- Costo total (USD) R{n} = Precio Unitario R{n} × (1 + Factor) ---
        for n in rondas:
            pu_l = L(col_precio_unitario(n))
            total_l = L(col_costo_total(n))
            if pu_l and total_l:
                if factor_l:
                    ws[f"{total_l}{r}"] = (
                        f'=IF({pu_l}{r}="","",{pu_l}{r}*(1+N({factor_l}{r})))'
                    )
                else:
                    # Sin tabla de Factor Logístico: el costo total es el precio unitario.
                    ws[f"{total_l}{r}"] = f'=IF({pu_l}{r}="","",{pu_l}{r})'

        # --- Homologación y Puntaje Técnico por Clave (ID|Proveedor) ---
        if tecnica_rango and id_l and prov_l:
            t_ini, t_fin = tecnica_rango
            clave = f'{id_l}{r}&"|"&{prov_l}{r}'
            if homolog_l:
                ws[f"{homolog_l}{r}"] = (
                    f"=IFERROR(INDEX('Detalle Evaluación Técnica'!$F${t_ini}:$F${t_fin},"
                    f"MATCH({clave},'Detalle Evaluación Técnica'!$A${t_ini}:$A${t_fin},0)),"
                    f'"S/DATOS")'
                )
            if puntaje_l:
                ws[f"{puntaje_l}{r}"] = (
                    f"=IFERROR(INDEX('Detalle Evaluación Técnica'!$D${t_ini}:$D${t_fin},"
                    f"MATCH({clave},'Detalle Evaluación Técnica'!$A${t_ini}:$A${t_fin},0)),"
                    f'"")'
                )

        # --- Precio Histórico por ID contra 'Evaluación Económica' ---
        if eval_economica_rango and id_l and hist_l:
            e_ini, e_fin = eval_economica_rango
            ws[f"{hist_l}{r}"] = (
                f"=IFERROR(INDEX('Evaluación Económica'!$D${e_ini}:$D${e_fin},"
                f"MATCH({id_l}{r},'Evaluación Económica'!$A${e_ini}:$A${e_fin},0)),"
                f'"")'
            )

        # --- Precio Histórico + factor logístico ---
        if hist_factor_l and hist_l and factor_l:
            ws[f"{hist_factor_l}{r}"] = (
                f'=IF(OR({hist_l}{r}="",NOT(ISNUMBER({hist_l}{r}))),"",'
                f'{hist_l}{r}*(1+N({factor_l}{r})))'
            )

        # --- Mínimo costo ofertado (por ID SAP, sobre la ronda vigente) ---
        if min_l and total_vigente_l and id_l:
            ws[f"{min_l}{r}"] = (
                f'=IFERROR(_xlfn.MINIFS({rng(total_vigente_l)},{rng(total_vigente_l)},">0",'
                f'{rng(id_l)},{id_l}{r}),"")'
            )

        # --- Desviación con Mín [%] ---
        if desv_l and total_vigente_l and min_l:
            ws[f"{desv_l}{r}"] = (
                f'=IF(OR({total_vigente_l}{r}="",{min_l}{r}="",{min_l}{r}=0),"",'
                f'({total_vigente_l}{r}/{min_l}{r})-1)'
            )

        # --- Baseline ---
        base_hist_l = hist_factor_l or hist_l
        if baseline_l and base_hist_l:
            infl_term = f'*(1+N({infl_l}{r}))' if infl_l else ""
            fallback = f'{total_base_l}{r}' if total_base_l else '""'
            ws[f"{baseline_l}{r}"] = (
                f'=IF(AND(ISNUMBER({base_hist_l}{r}),{base_hist_l}{r}>0),'
                f'{base_hist_l}{r}{infl_term},{fallback})'
            )

        # --- Escenarios ---
        _escribir_escenario(
            ws, r, tipo="auto_minimo",
            pct_l=L("% Adj. E1"), vol_l=L("Vol. E1"),
            monto_l=L("Monto E1\n(USD)"), ahorro_l=L("Ahorro_E1"),
            total_l=total_vigente_l, min_l=min_l, cantidad_l=cantidad_l,
            baseline_l=baseline_l,
        )
        _escribir_escenario(
            ws, r, tipo="auto_mejor_puntaje",
            pct_l=L("% Adj. E2"), vol_l=L("Vol. E2"),
            monto_l=L("Monto E2\n(USD)"), ahorro_l=L("Ahorro_E2"),
            total_l=total_vigente_l, min_l=min_l, cantidad_l=cantidad_l,
            baseline_l=baseline_l,
            homolog_l=homolog_l, puntaje_l=puntaje_l, id_l=id_l, rng=rng,
        )
        _escribir_escenario(
            ws, r, tipo="manual",
            pct_l=L("% Adj. E3\n← MANUAL"), vol_l=L("Vol. E3"),
            monto_l=L("Monto E3\n(USD)"), ahorro_l=L("Ahorro_E3"),
            total_l=total_vigente_l, min_l=min_l, cantidad_l=cantidad_l,
            baseline_l=baseline_l,
        )


def _escribir_escenario(ws, r, tipo, pct_l, vol_l, monto_l, ahorro_l,
                         total_l, min_l, cantidad_l, baseline_l,
                         homolog_l=None, puntaje_l=None, id_l=None, rng=None):
    if pct_l:
        if tipo == "auto_minimo" and total_l and min_l:
            ws[f"{pct_l}{r}"] = (
                f'=IF(OR({total_l}{r}="",{min_l}{r}=""),0,'
                f'IF(ROUND({total_l}{r},4)=ROUND({min_l}{r},4),1,0))'
            )
        elif tipo == "auto_mejor_puntaje" and total_l and puntaje_l and id_l and rng:
            # Mayor puntaje técnico del ítem, sea alto, medio o bajo.
            max_punt = (
                f'_xlfn.MAXIFS({rng(puntaje_l)},{rng(id_l)},{id_l}{r})'
            )
            # Desempate: entre los que igualan el puntaje máximo, el más barato.
            min_costo_max = (
                f'IFERROR(_xlfn.MINIFS({rng(total_l)},{rng(total_l)},">0",'
                f'{rng(id_l)},{id_l}{r},{rng(puntaje_l)},{max_punt}),"")'
            )
            ws[f"{pct_l}{r}"] = (
                f'=IF(OR({total_l}{r}="",{puntaje_l}{r}="",N({puntaje_l}{r})=0),0,'
                f'IF(AND(ROUND(N({puntaje_l}{r}),4)=ROUND({max_punt},4),'
                f'ROUND({total_l}{r},4)=ROUND({min_costo_max},4)),1,0))'
            )
        elif tipo == "manual":
            if ws[f"{pct_l}{r}"].value in (None, ""):
                ws[f"{pct_l}{r}"] = 0

    if vol_l and pct_l and cantidad_l:
        ws[f"{vol_l}{r}"] = (
            f'=IF(OR({pct_l}{r}="",{pct_l}{r}=0),"",{pct_l}{r}*{cantidad_l}{r})'
        )
    if monto_l and vol_l and total_l:
        ws[f"{monto_l}{r}"] = f'=IF({vol_l}{r}="","",{vol_l}{r}*${total_l}{r})'
    if ahorro_l and vol_l and total_l and baseline_l:
        ws[f"{ahorro_l}{r}"] = (
            f'=IF(OR(${baseline_l}{r}="",{vol_l}{r}=""),"",'
            f'(${baseline_l}{r}-${total_l}{r})*{vol_l}{r})'
        )
