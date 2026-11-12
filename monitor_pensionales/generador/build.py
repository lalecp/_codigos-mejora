import datetime as dt
import sys
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment

import importlib, os
_d = importlib.import_module(os.environ.get('DATOS', 'datos'))
POS, EST, RSK, EQUIV, HDR_P, HDR_E, HDR_R = _d.POS, _d.EST, _d.RSK, _d.EQUIV, _d.HDR_P, _d.HDR_E, _d.HDR_R

OUT = sys.argv[1]
N_POS = int(sys.argv[2]) if len(sys.argv) > 2 else 5000
N_ACT, N_PE, N_STR, N_RSK, N_EMI, N_PORT, N_DIM, N_QD, N_EQ = 3000, 2000, 2000, 3000, 300, 60, 40, 300, 200
R0 = 5

S_D, S_P, S_E, S_R, S_MA = '00_DASHBOARD', '01_POSICIONES_INPUT', '02_ESTRATEGIA_INPUT', '03_RIESGO_INPUT', '04_MAESTROS'
S_M, S_AC, S_RC, S_Q, S_I, S_C = '05_MODELO_POSICIONES', '06_ACTIVOS_COMUNES', '07_RIESGO_CALCULADO', '08_CALIDAD_DATOS', '09_INSTRUCCIONES', '10_CALCULOS'


def q(s):
    return f"'{s}'!"


def rg(s, col, n, r0=R0):
    return f"{q(s)}${col}${r0}:${col}${r0 + n - 1}"


def cell(s, col, r):
    return f"{q(s)}${col}${r}"


# ---------- estilos ----------
FN = 'Arial'
NAVY, TEAL, LBLUE, SOFTY = '1F3864', '2E75B6', 'DDEBF7', 'FFF2CC'
f_norm = Font(name=FN, size=9)
f_bold = Font(name=FN, size=9, bold=True)
f_in = Font(name=FN, size=9, color='0000FF')
f_hdr = Font(name=FN, size=9, bold=True, color='FFFFFF')
fill_hdr = PatternFill('solid', fgColor=NAVY)
fill_teal = PatternFill('solid', fgColor=TEAL)
fill_lb = PatternFill('solid', fgColor=LBLUE)
fill_in = PatternFill('solid', fgColor=SOFTY)
thin = Side(style='thin', color='BFBFBF')
b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
b_bot = Border(bottom=thin)
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = 'yyyy-mm-dd'

wb = Workbook()
ws_d = wb.active
ws_d.title = S_D
ws_p, ws_e, ws_r, ws_ma, ws_m, ws_ac, ws_rc, ws_q, ws_i, ws_c = [wb.create_sheet(n) for n in
    (S_P, S_E, S_R, S_MA, S_M, S_AC, S_RC, S_Q, S_I, S_C)]


def add_table(ws, name, ref, style='TableStyleMedium2'):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)
    return t


def hdr_row(ws, r, c0, names, widths=None):
    for i, n in enumerate(names):
        c = ws.cell(r, c0 + i, n)
        c.font, c.fill = f_hdr, fill_hdr
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if widths:
            ws.column_dimensions[L(c0 + i)].width = widths[i]


def title(ws, text, sub=None):
    ws['A1'] = text
    ws['A1'].font = Font(name=FN, size=13, bold=True, color=NAVY)
    if sub:
        ws['A2'] = sub
        ws['A2'].font = Font(name=FN, size=9, italic=True, color='595959')


# =====================================================================
# 01 / 02 / 03  ENTRADAS (encabezado en fila 1, datos desde fila 2)
# =====================================================================
def write_input(ws, hdr, rows, fmts, widths, tname):
    for j, h in enumerate(hdr, 1):
        ws.cell(1, j, h)
        ws.column_dimensions[L(j)].width = widths[j - 1]
    for i, row in enumerate(rows, 2):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.font = f_in
            if fmts.get(j):
                c.number_format = fmts[j]
    add_table(ws, tname, f"A1:{L(len(hdr))}{len(rows) + 1}")
    for j in range(1, len(hdr) + 1):
        ws.cell(1, j).font = Font(name=FN, size=9, bold=True, color='FFFFFF')
        ws.cell(1, j).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = 'A2'
    ws.sheet_properties.tabColor = 'FFC000'


write_input(ws_p, HDR_P, POS, {1: DATE, 14: DATE, 15: NUM2, 16: NUM2, 17: '0.00', 18: '0.00%'},
            [11, 11, 34, 12, 15, 16, 13, 28, 18, 24, 26, 12, 24, 12, 14, 14, 9, 9, 8, 16], 'tblPosiciones')
ws_p['O1'].comment = Comment('Unidad del ejemplo: COP MM. Si la sábana real viene en pesos, cambie el factor en 04_MAESTROS!B8.', 'Monitor')

# Estrategia: G = DireccionEstrategia (fórmula de la tabla; el modelo NO depende de ella)
for i, row in enumerate(EST, 2):
    pass
est_rows = [list(r[:6]) + [None, r[6]] for r in EST]
write_input(ws_e, HDR_E, est_rows, {1: DATE, 5: '0.00;(0.00);0.00'}, [11, 11, 16, 30, 12, 18, 16, 60], 'tblEstrategia')
for i in range(2, len(EST) + 2):
    c = ws_e.cell(i, 7, f'=IF(E{i}="","Sin definir",IF(E{i}>0,"Aumentar",IF(E{i}<0,"Reducir","Mantener")))')
    c.font = f_norm
ws_e['E1'].comment = Comment('Positivo = aumentar / alargar. Negativo = reducir / acortar. 0 = mantener. Vacío = sin estrategia definida (no es cero).', 'Monitor')

write_input(ws_r, HDR_R, RSK, {1: DATE, 6: '0.0%', 7: NUM2, 8: NUM2}, [11, 11, 14, 30, 22, 12, 13, 13, 11, 22], 'tblRiesgo')
ws_r['F1'].comment = Comment('Fracción del valor total del portafolio (0.10 = 10%). Si llega como 10, el modelo lo interpreta como 10%.', 'Monitor')
ws_r['G1'].comment = Comment('Misma unidad que ValorPresente. Si existe, tiene prioridad sobre el porcentaje.', 'Monitor')

# =====================================================================
# 04 MAESTROS
# =====================================================================
title(ws_ma, 'Maestros y parámetros', 'Celdas con fondo amarillo y letra azul = editables. El resto son fórmulas.')
params = [
    ('Umbral Seguimiento (utilización >=)', 0.75, 'Desde este nivel el emisor pasa a Seguimiento'),
    ('Umbral Alerta (utilización >=)', 0.85, 'Desde este nivel el emisor pasa a Alerta'),
    ('Umbral Al límite o excedido (utilización >=)', 1.00, 'Desde este nivel se considera al límite o excedido'),
    ('Longitud de ISIN válido (0 = no validar)', 12, 'Si el ISIN no tiene esta longitud se usa la clave alternativa'),
    ('Factor ValorPresente a COP MM', 1, 'Ejemplo en COP MM = 1. Si la sábana viene en pesos use 0.000001'),
    ('Modo de datos', 'Sintético', 'Sintético muestra el aviso en el dashboard. Cambie a Real con datos reales'),
    ('TipoPortafolio incluido en el modelo', 'Pensional', 'Solo entran al modelo las filas con este TipoPortafolio'),
]
hdr_row(ws_ma, 3, 1, ['Parámetro', 'Valor', 'Nota'], [42, 14, 62])
for i, (a, b, c) in enumerate(params, 4):
    ws_ma.cell(i, 1, a).font = f_norm
    x = ws_ma.cell(i, 2, b)
    x.font, x.fill, x.border = f_in, fill_in, b_all
    ws_ma.cell(i, 3, c).font = Font(name=FN, size=8, italic=True, color='595959')
for r_ in (4, 5, 6):
    ws_ma.cell(r_, 2).number_format = '0%'
ws_ma['A11'], ws_ma['A12'] = 'Índice de la dimensión de estrategia (fórmula)', 'Dimensión de estrategia activa (fórmula)'
ws_ma['B11'] = f'=IFERROR(MATCH(B12,E5:E10,0),1)'
ws_ma['B12'] = f"={q(S_D)}$P$8"
ws_ma['C11'] = 'Si la dimensión elegida no es válida se usa Activo (primera de la lista)'
ws_ma['A13'], ws_ma['B13'] = 'Capacidad del modelo (filas de posiciones)', N_POS
for r_ in (11, 12, 13):
    ws_ma.cell(r_, 1).font = f_norm
    ws_ma.cell(r_, 2).font = f_norm
    ws_ma.cell(r_, 3).font = Font(name=FN, size=8, italic=True, color='595959')
P_U1, P_U2, P_U3 = cell(S_MA, 'B', 4), cell(S_MA, 'B', 5), cell(S_MA, 'B', 6)
P_LEN, P_FAC, P_MODO, P_TIPO = cell(S_MA, 'B', 7), cell(S_MA, 'B', 8), cell(S_MA, 'B', 9), cell(S_MA, 'B', 10)
P_DIMI, P_DIMN = cell(S_MA, 'B', 11), cell(S_MA, 'B', 12)
dv = DataValidation(type='list', formula1='"Sintético,Real"', allow_blank=False)
ws_ma.add_data_validation(dv); dv.add('B9')

DIMS = ['Activo', 'Nodo', 'Indicador', 'TipoPapel', 'Sector', 'Emisor']
hdr_row(ws_ma, 4, 5, ['DimensionEstrategia'], [20])
for i, d in enumerate(DIMS + ['Duracion'], 5):
    ws_ma.cell(i, 5, d).font = f_norm
ws_ma['E12'] = 'Duracion = dirección de duración del portafolio (alargar / acortar)'
ws_ma['E12'].font = Font(name=FN, size=8, italic=True, color='595959')
hdr_row(ws_ma, 4, 7, ['EstadoEstrategia'], [18])
for i, d in enumerate(['Aumentar', 'Reducir', 'Mantener', 'Sin definir'], 5):
    ws_ma.cell(i, 7, d).font = f_norm
NODOS = ['Liquidez / menor a 1 año', '1-3 años', '3-5 años', '5-7 años', '7-10 años', '10-15 años', 'Mayor a 15 años', 'Sin clasificar']
hdr_row(ws_ma, 4, 9, ['Orden', 'Nodo'], [7, 26])
for i, d in enumerate(NODOS, 5):
    ws_ma.cell(i, 9, i - 4).font = f_in
    c = ws_ma.cell(i, 10, d); c.font, c.fill = f_in, fill_in
RG_NODOS = f"{q(S_MA)}$J$5:$J$12"
# catálogo de portafolios (automático)
hdr_row(ws_ma, 4, 12, ['k', 'CodigoPortafolio', 'NombrePortafolio'], [5, 17, 38])
# equivalencias NIT
hdr_row(ws_ma, 4, 16, ['NITOrigen', 'NITNormalizado'], [16, 16])
for i in range(N_EQ):
    for j in (16, 17):
        c = ws_ma.cell(5 + i, j); c.font, c.fill = f_in, fill_in
        c.number_format = '@'
for i, (a, b) in enumerate(EQUIV, 5):
    ws_ma.cell(i, 16, a); ws_ma.cell(i, 17, b)
RG_EQO, RG_EQN = rg(S_MA, 'P', N_EQ), rg(S_MA, 'Q', N_EQ)
ws_ma['P3'] = 'Equivalencias de NIT (editable): se aplican a posiciones y a riesgo'
ws_ma['P3'].font = f_bold
# catálogos de valores válidos
CAT_SECTOR = ['Gobierno nacional', 'Financiero', 'Energía', 'Infraestructura', 'Servicios públicos', 'Corporativo', 'Liquidez', 'Otros', 'Sin clasificar']
CAT_PAPEL = ['TES', 'Bono ordinario', 'Bono subordinado', 'CDT', 'Titularización', 'FIC', 'Cuenta bancaria / liquidez', 'Otro', 'Sin clasificar']
CAT_IND = ['Tasa fija', 'IPC', 'UVR', 'IBR', 'DTF', 'Liquidez', 'Otro', 'Sin clasificar']
hdr_row(ws_ma, 4, 19, ['Sector', 'TipoPapel', 'Indicador'], [20, 26, 14])
ws_ma['S3'] = 'Catálogos de clasificación (editables, 40 filas)'
ws_ma['S3'].font = f_bold
for j, lst in ((19, CAT_SECTOR), (20, CAT_PAPEL), (21, CAT_IND)):
    for i in range(N_DIM):
        c = ws_ma.cell(5 + i, j, lst[i] if i < len(lst) else None); c.font, c.fill = f_in, fill_in
RG_CSEC, RG_CPAP, RG_CIND = rg(S_MA, 'S', N_DIM), rg(S_MA, 'T', N_DIM), rg(S_MA, 'U', N_DIM)
# catálogo de emisores (automático)
hdr_row(ws_ma, 4, 23, ['k', 'NITEmisor', 'Emisor', 'Sector'], [5, 14, 30, 20])
ws_ma['L3'] = 'Catálogo de portafolios (automático desde la sábana)'; ws_ma['L3'].font = f_bold
ws_ma['W3'] = 'Catálogo de emisores (automático desde la sábana)'; ws_ma['W3'].font = f_bold
ws_ma.freeze_panes = 'A5'

# =====================================================================
# 05 MODELO
# =====================================================================
MCOLS = ['FilaOrigen', 'Activa', 'Incluido', 'FechaCorte', 'CodigoPortafolio', 'NombrePortafolio', 'ISIN', 'Nemotecnico',
         'NITEmisor', 'Emisor', 'Sector', 'TipoActivo', 'TipoPapel', 'Indicador', 'Nodo', 'FechaVencimiento', 'ValorPresente',
         'Duracion', 'Moneda', 'ClaveActivo', 'ClavePortafolioActivo', 'ClavePortafolioEmisor', 'ValorTotalPortafolio',
         'PesoPortafolio', 'CategoriaEstrategiaAplicable', 'FilaEstrategia', 'ValorEstrategia', 'DireccionEstrategia',
         'ExposicionEmisor', 'FilaRiesgo', 'LimiteAplicable', 'UtilizacionLimite', 'HolguraLimite', 'EstadoRiesgo',
         'Seleccionado', 'PrimPA', 'SelPA', 'SelAumentar', 'SelReducir', 'SelMantener', 'SelSinDefinir', 'NumSelActivo',
         'RankPortafolio', 'RankActivo', 'RankPortEmisor', 'RankTipoActivo', 'RankSector', 'RankEmisor', 'RankIndicador',
         'RankTipoPapel', 'VPxDur', 'VPconDur']
QFLAGS = [  # (columna, etiqueta, severidad)
    ('qCodVacio', 'CodigoPortafolio vacío', 'Error'),
    ('qNomVacio', 'NombrePortafolio vacío', 'Error'),
    ('qNITVacio', 'NITEmisor vacío', 'Error'),
    ('qVPNoNum', 'ValorPresente no numérico', 'Error'),
    ('qNoPensional', 'TipoPortafolio diferente de Pensional (excluido)', 'Informativo'),
    ('qISINAlt', 'ISIN vacío con clave alternativa disponible', 'Advertencia'),
    ('qISINSinDatos', 'ISIN vacío sin datos suficientes para clave alternativa', 'Error'),
    ('qSectorVacio', 'Sector vacío', 'Advertencia'),
    ('qIndicadorVacio', 'Indicador vacío', 'Advertencia'),
    ('qNodoVacio', 'Nodo vacío', 'Advertencia'),
    ('qEstrSinCoinc', 'Estrategia sin coincidencia (dimensión activa)', 'Advertencia'),
    ('qSinLimite', 'Emisor sin límite de riesgo', 'Advertencia'),
    ('qPortNoRiesgo', 'Código de portafolio no encontrado en riesgo', 'Advertencia'),
    ('qLimiteCero', 'Límite igual a cero con exposición positiva', 'Error crítico'),
    ('qClaveContra', 'ClaveActivo duplicada con atributos contradictorios', 'Error crítico'),
    ('qISINLong', 'ISIN con longitud no válida (se usa clave alternativa)', 'Advertencia'),
    ('qFueraCatalogo', 'Nodo, sector, indicador o tipo de papel fuera de catálogo', 'Advertencia'),
    ('qVPNegativo', 'ValorPresente negativo', 'Informativo'),
]
MCOLS += [f[0] for f in QFLAGS] + ['EstadoCalidad', 'DetalleCalidad', 'RankIncidencia']
mc = {n: L(i + 1) for i, n in enumerate(MCOLS)}


def MR(n):
    return rg(S_M, mc[n], N_POS)


# rangos de ayudas en 10_CALCULOS (se definen antes porque el modelo los usa)
cc = {}
_cur = [1]


def alloc(block, names):
    for n in names:
        cc[(block, n)] = L(_cur[0]); _cur[0] += 1
    _cur[0] += 1


alloc('S', ['fila', 'Cod', 'Dim', 'Cat', 'Val', 'Unidad', 'Dir', 'Key', 'SortKey', 'Rank', 'Dup', 'DimInv'])
alloc('R', ['fila', 'Cod', 'NIT', 'Key', 'Pct', 'Monto', 'TipoLimite', 'Calif', 'Emisor', 'Dup'])
for b in ('TA', 'SE', 'IN', 'TP'):
    alloc(b, ['k', 'Cat', 'Val', 'Rank'])
    alloc(b + 't', ['pos', 'Label', 'Peso'])
alloc('EM', ['k', 'NIT', 'Emisor', 'Val', 'Rank'])
alloc('EMt', ['pos', 'Label', 'Peso'])
alloc('NO', ['orden', 'Label', 'Val', 'Peso'])


def CR(block, n, cap):
    return rg(S_C, cc[(block, n)], cap)


C_TOT, C_SEL, C_TOTSEL = cell(S_C, 'B', 1), cell(S_C, 'D', 1), cell(S_C, 'F', 1)
SLOTS = ['B', 'F', 'J', 'N', 'R', 'V']
RG_SLOTS = f"{q(S_D)}$B$11:$Y$11"

title(ws_m, 'Modelo de posiciones (normalización y cruces). No editar: todo son fórmulas.',
      'Cada fila lee una fila de 01_POSICIONES_INPUT (columna FilaOrigen). Capacidad: %d filas.' % N_POS)
hdr_row(ws_m, 4, 1, MCOLS)
for i, n in enumerate(MCOLS, 1):
    ws_m.column_dimensions[L(i)].width = 14
ws_m.row_dimensions[4].height = 30

FLAG_SEV = {f[0]: f[2] for f in QFLAGS}


def sev_sum(sev, r):
    cols = [f"{mc[f]}{r}" for f, _, s in QFLAGS if s == sev]
    return '(' + '+'.join(cols) + ')'


def model_row(r):
    m = {k: f"{v}{r}" for k, v in mc.items()}
    P = q(S_P)

    def raw(c):
        return f"INDEX({P}${c}:${c},$A{r})"

    def txt(c):
        return f'TRIM({raw(c)}&"")'

    def first(col, cond_extra=''):
        cl = mc[col]
        me = None
        return cl

    def rank_distinct(self_col, src):
        s, cl = mc[src], mc[self_col]
        return (f'=IF(AND({m["Incluido"]}=1,{s}{r}<>""),IF(COUNTIF({s}${R0}:{s}{r},{s}{r})=1,'
                f'MAX({cl}${R0 - 1}:{cl}{r - 1})+1,""),"")')

    inc = m['Incluido']
    f = {}
    f['FilaOrigen'] = r - R0 + 2
    f['Activa'] = f'=IF(COUNTA(INDEX({P}$A:$T,$A{r},0))>0,1,0)'
    f['Incluido'] = f'=IF({m["Activa"]}=1,IF({txt("D")}={P_TIPO},1,0),0)'
    f['FechaCorte'] = f'=IF({inc}=1,IF(ISNUMBER({raw("A")}),{raw("A")},""),"")'
    f['CodigoPortafolio'] = f'=IF({inc}=1,{txt("B")},"")'
    f['NombrePortafolio'] = f'=IF({inc}=1,{txt("C")},"")'
    f['ISIN'] = f'=IF({inc}=1,{txt("E")},"")'
    f['Nemotecnico'] = f'=IF({inc}=1,{txt("F")},"")'
    f['NITEmisor'] = f'=IF({inc}=1,IFERROR(INDEX({RG_EQN},MATCH({txt("G")},{RG_EQO},0))&"",{txt("G")}),"")'
    f['Emisor'] = f'=IF({inc}=1,{txt("H")},"")'
    for n, c in (('Sector', 'I'), ('TipoActivo', 'J'), ('TipoPapel', 'K'), ('Indicador', 'L'), ('Nodo', 'M')):
        f[n] = f'=IF({inc}=1,IF({txt(c)}="","Sin clasificar",{txt(c)}),"")'
    f['FechaVencimiento'] = f'=IF({inc}=1,IF({raw("N")}="","",{raw("N")}),"")'
    f['ValorPresente'] = f'=IF({inc}=1,IF(ISNUMBER({raw("O")}),{raw("O")},0),"")'
    f['Duracion'] = f'=IF({inc}=1,IF(ISNUMBER({raw("Q")}),{raw("Q")},""),"")'
    f['Moneda'] = f'=IF({inc}=1,{txt("S")},"")'
    fv = m['FechaVencimiento']
    f['ClaveActivo'] = (f'=IF({inc}<>1,"",IF(AND({m["ISIN"]}<>"",OR({P_LEN}=0,LEN({m["ISIN"]})={P_LEN})),{m["ISIN"]},'
                        f'IF({fv}<>"",{m["NITEmisor"]}&"|"&{m["Nemotecnico"]}&"|"&IF(ISNUMBER({fv}),YEAR({fv})*10000+MONTH({fv})*100+DAY({fv}),{fv})&"|"&{m["Moneda"]},'
                        f'{m["NITEmisor"]}&"|"&{m["Nemotecnico"]}&"|"&{m["TipoPapel"]}&"|"&{m["Indicador"]})))')
    f['ClavePortafolioActivo'] = f'=IF({inc}=1,{m["CodigoPortafolio"]}&"|"&{m["ClaveActivo"]},"")'
    f['ClavePortafolioEmisor'] = f'=IF({inc}=1,{m["CodigoPortafolio"]}&"|"&{m["NITEmisor"]},"")'
    f['ValorTotalPortafolio'] = f'=IF({inc}=1,SUMIFS({MR("ValorPresente")},{MR("CodigoPortafolio")},{m["CodigoPortafolio"]}),"")'
    f['PesoPortafolio'] = f'=IF({inc}=1,IF({m["ValorTotalPortafolio"]}=0,"",{m["ValorPresente"]}/{m["ValorTotalPortafolio"]}),"")'
    f['CategoriaEstrategiaAplicable'] = (f'=IF({inc}=1,CHOOSE({P_DIMI},{m["TipoActivo"]},{m["Nodo"]},{m["Indicador"]},'
                                         f'{m["TipoPapel"]},{m["Sector"]},{m["Emisor"]}),"")')
    f['FilaEstrategia'] = (f'=IF({inc}=1,IFERROR(MATCH({m["CodigoPortafolio"]}&"|"&INDEX({rg(S_MA, "E", 6)},{P_DIMI})&"|"&{m["CategoriaEstrategiaAplicable"]},'
                           f'{CR("S", "Key", N_STR)},0),0),0)')
    f['ValorEstrategia'] = f'=IF({m["FilaEstrategia"]}=0,"",INDEX({CR("S", "Val", N_STR)},{m["FilaEstrategia"]}))'
    f['DireccionEstrategia'] = f'=IF({inc}<>1,"",IF({m["FilaEstrategia"]}=0,"Sin definir",INDEX({CR("S", "Dir", N_STR)},{m["FilaEstrategia"]})))'
    f['ExposicionEmisor'] = f'=IF({inc}=1,SUMIFS({MR("ValorPresente")},{MR("ClavePortafolioEmisor")},{m["ClavePortafolioEmisor"]}),"")'
    f['FilaRiesgo'] = f'=IF({inc}=1,IFERROR(MATCH({m["ClavePortafolioEmisor"]},{CR("R", "Key", N_RSK)},0),0),0)'
    fr = m['FilaRiesgo']
    mon, pct = f'INDEX({CR("R", "Monto", N_RSK)},{fr})', f'INDEX({CR("R", "Pct", N_RSK)},{fr})'
    f['LimiteAplicable'] = f'=IF({fr}=0,"",IF(ISNUMBER({mon}),{mon},IF(ISNUMBER({pct}),{m["ValorTotalPortafolio"]}*{pct},"")))'
    la, ex = m['LimiteAplicable'], m['ExposicionEmisor']
    f['UtilizacionLimite'] = f'=IF(ISNUMBER({la}),IF({la}=0,"",{ex}/{la}),"")'
    f['HolguraLimite'] = f'=IF(ISNUMBER({la}),{la}-{ex},"")'
    ut = m['UtilizacionLimite']
    f['EstadoRiesgo'] = (f'=IF({inc}<>1,"",IF({fr}=0,"Sin información de riesgo",IF(NOT(ISNUMBER({la})),"Sin límite informado",'
                         f'IF({la}=0,IF({ex}>0,"Límite cero","Dentro de límite"),IF({ut}>={P_U3},"Al límite o excedido",'
                         f'IF({ut}>={P_U2},"Alerta",IF({ut}>={P_U1},"Seguimiento","Dentro de límite")))))))')
    f['Seleccionado'] = f'=IF({inc}=1,IF(COUNTIF({RG_SLOTS},{m["CodigoPortafolio"]})>0,1,0),0)'
    u = mc['ClavePortafolioActivo']
    f['PrimPA'] = f'=IF({inc}=1,IF(COUNTIF({u}${R0}:{u}{r},{u}{r})=1,1,0),0)'
    f['SelPA'] = f'={m["Seleccionado"]}*{m["PrimPA"]}'
    for n, w in (('SelAumentar', 'Aumentar'), ('SelReducir', 'Reducir'), ('SelMantener', 'Mantener'), ('SelSinDefinir', 'Sin definir')):
        f[n] = f'=IF({m["SelPA"]}=1,IF({m["DireccionEstrategia"]}="{w}",1,0),0)'
    f['NumSelActivo'] = f'=IF({m["SelPA"]}=1,SUMIFS({MR("SelPA")},{MR("ClaveActivo")},{m["ClaveActivo"]}),0)'
    f['RankPortafolio'] = rank_distinct('RankPortafolio', 'CodigoPortafolio')
    f['RankActivo'] = rank_distinct('RankActivo', 'ClaveActivo')
    f['RankPortEmisor'] = rank_distinct('RankPortEmisor', 'ClavePortafolioEmisor')
    f['RankTipoActivo'] = rank_distinct('RankTipoActivo', 'TipoActivo')
    f['RankSector'] = rank_distinct('RankSector', 'Sector')
    f['RankEmisor'] = rank_distinct('RankEmisor', 'NITEmisor')
    f['RankIndicador'] = rank_distinct('RankIndicador', 'Indicador')
    f['RankTipoPapel'] = rank_distinct('RankTipoPapel', 'TipoPapel')
    f['VPxDur'] = f'=IF({inc}=1,IF(ISNUMBER({m["Duracion"]}),{m["ValorPresente"]}*{m["Duracion"]},0),0)'
    f['VPconDur'] = f'=IF({inc}=1,IF(ISNUMBER({m["Duracion"]}),{m["ValorPresente"]},0),0)'
    act = m['Activa']
    conds = {
        'qCodVacio': f'AND({act}=1,{txt("B")}="")',
        'qNomVacio': f'AND({act}=1,{txt("C")}="")',
        'qNITVacio': f'AND({act}=1,{txt("G")}="")',
        'qVPNoNum': f'AND({act}=1,NOT(ISNUMBER({raw("O")})))',
        'qNoPensional': f'AND({act}=1,{inc}=0)',
        'qISINAlt': f'AND({inc}=1,{m["ISIN"]}="",{m["NITEmisor"]}<>"",{m["Nemotecnico"]}<>"")',
        'qISINSinDatos': f'AND({inc}=1,{m["ISIN"]}="",OR({m["NITEmisor"]}="",{m["Nemotecnico"]}=""))',
        'qSectorVacio': f'AND({inc}=1,{txt("I")}="")',
        'qIndicadorVacio': f'AND({inc}=1,{txt("L")}="")',
        'qNodoVacio': f'AND({inc}=1,{txt("M")}="")',
        'qEstrSinCoinc': f'AND({inc}=1,{m["FilaEstrategia"]}=0)',
        'qSinLimite': f'AND({inc}=1,{m["qPortNoRiesgo"]}=0,NOT(ISNUMBER({la})))',
        'qPortNoRiesgo': f'AND({inc}=1,{m["CodigoPortafolio"]}<>"",COUNTIF({CR("R", "Cod", N_RSK)},{m["CodigoPortafolio"]})=0)',
        'qLimiteCero': f'AND({inc}=1,{m["EstadoRiesgo"]}="Límite cero")',
        'qClaveContra': (f'AND({inc}=1,{m["ClaveActivo"]}<>"",COUNTIF({MR("ClaveActivo")},{m["ClaveActivo"]})<>'
                         f'COUNTIFS({MR("ClaveActivo")},{m["ClaveActivo"]},{MR("NITEmisor")},{m["NITEmisor"]},'
                         f'{MR("TipoPapel")},{m["TipoPapel"]},{MR("Indicador")},{m["Indicador"]}))'),
        'qISINLong': f'AND({inc}=1,{m["ISIN"]}<>"",{P_LEN}>0,LEN({m["ISIN"]})<>{P_LEN})',
        'qFueraCatalogo': (f'AND({inc}=1,OR(COUNTIF({RG_NODOS},{m["Nodo"]})=0,COUNTIF({RG_CSEC},{m["Sector"]})=0,'
                           f'COUNTIF({RG_CIND},{m["Indicador"]})=0,COUNTIF({RG_CPAP},{m["TipoPapel"]})=0))'),
        'qVPNegativo': f'AND({inc}=1,ISNUMBER({raw("O")}),{raw("O")}<0)',
    }
    for k, v in conds.items():
        f[k] = f'=IF({act}=0,0,IF({v},1,0))'
    f['EstadoCalidad'] = (f'=IF({act}=0,"",IF({sev_sum("Error crítico", r)}>0,"Error crítico",IF({sev_sum("Error", r)}>0,"Error",'
                          f'IF({sev_sum("Advertencia", r)}>0,"Advertencia",IF({sev_sum("Informativo", r)}>0,"Informativo","OK")))))')
    parts = '&'.join(f'IF({mc[n]}{r}=1,"{lab}; ","")' for n, lab, _ in QFLAGS)
    f['DetalleCalidad'] = f'=IF({act}=0,"",{parts})'
    ri = mc['RankIncidencia']
    f['RankIncidencia'] = f'=IF({act}=0,"",IF({m["EstadoCalidad"]}="OK","",MAX({ri}${R0 - 1}:{ri}{r - 1})+1))'
    return f


MFMT = {'FechaCorte': DATE, 'FechaVencimiento': DATE, 'ValorPresente': NUM2, 'ValorTotalPortafolio': NUM2, 'PesoPortafolio': '0.00%',
        'ExposicionEmisor': NUM2, 'LimiteAplicable': NUM2, 'UtilizacionLimite': '0.0%', 'HolguraLimite': NUM2, 'Duracion': '0.00',
        'VPxDur': NUM2, 'VPconDur': NUM2}
for r in range(R0, R0 + N_POS):
    f = model_row(r)
    for j, n in enumerate(MCOLS, 1):
        c = ws_m.cell(r, j, f[n])
        if n in MFMT:
            c.number_format = MFMT[n]
add_table(ws_m, 'tblModeloPosiciones', f"A4:{L(len(MCOLS))}{R0 + N_POS - 1}", 'TableStyleLight1')
ws_m.freeze_panes = 'F5'

# =====================================================================
# 10 CALCULOS (ayudas)
# =====================================================================
title(ws_c, None)
ws_c['A1'], ws_c['C1'], ws_c['E1'] = 'TotalSel', 'Sel', 'TotalSeleccionados'
ws_c['B1'] = f'=SUMIFS({MR("ValorPresente")},{MR("CodigoPortafolio")},D1)'
ws_c['D1'] = f"={q(S_D)}$B$8&\"\""
ws_c['F1'] = f'=SUMIFS({MR("ValorPresente")},{MR("Seleccionado")},1)'
ws_c['A2'] = 'Hoja de ayudas del dashboard. No editar.'
ws_c['A2'].font = Font(name=FN, size=9, italic=True)


def chdr(block, names):
    for n in names:
        c = ws_c[f"{cc[(block, n)]}4"]
        c.value = f"{block}_{n}"; c.font, c.fill = f_hdr, fill_hdr


# estrategia normalizada
chdr('S', ['fila', 'Cod', 'Dim', 'Cat', 'Val', 'Unidad', 'Dir', 'Key', 'SortKey', 'Rank', 'Dup', 'DimInv'])
E_ = q(S_E)
for i in range(N_STR):
    r = R0 + i
    s = {n: f"{cc[('S', n)]}{r}" for n in ['fila', 'Cod', 'Dim', 'Cat', 'Val', 'Unidad', 'Dir', 'Key', 'SortKey', 'Rank', 'Dup', 'DimInv']}

    def er(c):
        return f"INDEX({E_}${c}:${c},{s['fila']})"
    ws_c[s['fila']] = i + 2
    ws_c[s['Cod']] = f'=TRIM({er("B")}&"")'
    ws_c[s['Dim']] = f'=TRIM({er("C")}&"")'
    ws_c[s['Cat']] = f'=TRIM({er("D")}&"")'
    ws_c[s['Val']] = f'=IF({s["Cod"]}="","",IF(ISNUMBER({er("E")}),{er("E")},""))'
    ws_c[s['Unidad']] = f'=TRIM({er("F")}&"")'
    ws_c[s['Dir']] = f'=IF({s["Cod"]}="","",IF(ISNUMBER({s["Val"]}),IF({s["Val"]}>0,"Aumentar",IF({s["Val"]}<0,"Reducir","Mantener")),"Sin definir"))'
    ws_c[s['Key']] = f'=IF({s["Cod"]}="","",{s["Cod"]}&"|"&{s["Dim"]}&"|"&{s["Cat"]})'
    ws_c[s['SortKey']] = (f'=IF(OR({s["Cod"]}="",{s["Cod"]}<>{C_SEL},AND({s["Dim"]}<>{P_DIMN},{s["Dim"]}<>"Duracion")),-1,IF({s["Dir"]}="Aumentar",3000000000+ROUND({s["Val"]},4),'
                          f'IF({s["Dir"]}="Reducir",2000000000-ROUND({s["Val"]},4),IF({s["Dir"]}="Mantener",1000000000,1))))')
    sk, fl = CR('S', 'SortKey', N_STR), CR('S', 'fila', N_STR)
    ws_c[s['Rank']] = f'=IF({s["SortKey"]}<0,"",SUMPRODUCT(({sk}>{s["SortKey"]})+0)+SUMPRODUCT(({sk}={s["SortKey"]})*({fl}<{s["fila"]}))+1)'
    ws_c[s['Dup']] = f'=IF({s["Key"]}="",0,IF(COUNTIF({CR("S", "Key", N_STR)},{s["Key"]})>1,1,0))'
    ws_c[s['DimInv']] = f'=IF({s["Cod"]}="",0,IF(COUNTIF({rg(S_MA, "E", 7)},{s["Dim"]})=0,1,0))'

# riesgo normalizado
RN = ['fila', 'Cod', 'NIT', 'Key', 'Pct', 'Monto', 'TipoLimite', 'Calif', 'Emisor', 'Dup']
chdr('R', RN)
R_ = q(S_R)
for i in range(N_RSK):
    r = R0 + i
    s = {n: f"{cc[('R', n)]}{r}" for n in RN}

    def rr(c):
        return f"INDEX({R_}${c}:${c},{s['fila']})"
    ws_c[s['fila']] = i + 2
    ws_c[s['Cod']] = f'=TRIM({rr("B")}&"")'
    ws_c[s['NIT']] = f'=IF({s["Cod"]}="","",IFERROR(INDEX({RG_EQN},MATCH(TRIM({rr("C")}&""),{RG_EQO},0))&"",TRIM({rr("C")}&"")))'
    ws_c[s['Key']] = f'=IF({s["Cod"]}="","",{s["Cod"]}&"|"&{s["NIT"]})'
    ws_c[s['Pct']] = f'=IF({s["Cod"]}="","",IF(ISNUMBER({rr("F")}),IF({rr("F")}>1,{rr("F")}/100,{rr("F")}),""))'
    ws_c[s['Monto']] = f'=IF({s["Cod"]}="","",IF(ISNUMBER({rr("G")}),{rr("G")},""))'
    ws_c[s['TipoLimite']] = f'=TRIM({rr("E")}&"")'
    ws_c[s['Calif']] = f'=TRIM({rr("I")}&"")'
    ws_c[s['Emisor']] = f'=TRIM({rr("D")}&"")'
    ws_c[s['Dup']] = f'=IF({s["Key"]}="",0,IF(COUNTIF({CR("R", "Key", N_RSK)},{s["Key"]})>1,1,0))'


def dim_block(b, src, rankcol, ntop, cap=N_DIM, emisor=False):
    names = ['k', 'NIT', 'Emisor', 'Val', 'Rank'] if emisor else ['k', 'Cat', 'Val', 'Rank']
    chdr(b, names); chdr(b + 't', ['pos', 'Label', 'Peso'])
    catn = 'NIT' if emisor else 'Cat'
    labn = 'Emisor' if emisor else 'Cat'
    for i in range(cap):
        r = R0 + i
        s = {n: f"{cc[(b, n)]}{r}" for n in names}
        ws_c[s['k']] = i + 1
        ws_c[s[catn]] = f'=IFERROR(INDEX({MR(src)},MATCH({s["k"]},{MR(rankcol)},0)),"")'
        if emisor:
            ws_c[s['Emisor']] = f'=IFERROR(INDEX({MR("Emisor")},MATCH({s["k"]},{MR(rankcol)},0)),"")'
        ws_c[s['Val']] = f'=IF({s[catn]}="",-1,SUMIFS({MR("ValorPresente")},{MR("CodigoPortafolio")},{C_SEL},{MR(src)},{s[catn]}))'
        vr, kr = CR(b, 'Val', cap), CR(b, 'k', cap)
        ws_c[s['Rank']] = f'=IF({s["Val"]}>0,SUMPRODUCT(({vr}>{s["Val"]})+0)+SUMPRODUCT(({vr}={s["Val"]})*({kr}<{s["k"]}))+1,"")'
    t = {n: cc[(b + 't', n)] for n in ['pos', 'Label', 'Peso']}
    for p in range(1, ntop):
        r = R0 + p - 1
        ws_c[f"{t['pos']}{r}"] = p
        ws_c[f"{t['Label']}{r}"] = f'=IFERROR(INDEX({CR(b, labn, cap)},MATCH({t["pos"]}{r},{CR(b, "Rank", cap)},0)),"")'
        ws_c[f"{t['Peso']}{r}"] = f'=IF(OR({t["Label"]}{r}="",{C_TOT}<=0),0,INDEX({CR(b, "Val", cap)},MATCH({t["pos"]}{r},{CR(b, "Rank", cap)},0))/{C_TOT})'
        ws_c[f"{t['Peso']}{r}"].number_format = '0.0%'
    r = R0 + ntop - 1
    ws_c[f"{t['pos']}{r}"] = ntop
    ws_c[f"{t['Peso']}{r}"] = f'=IF({C_TOT}<=0,0,MAX(0,ROUND(1-SUM({t["Peso"]}{R0}:{t["Peso"]}{r - 1}),6)))'
    ws_c[f"{t['Peso']}{r}"].number_format = '0.0%'
    ws_c[f"{t['Label']}{r}"] = f'=IF({t["Peso"]}{r}>0.00005,"Otros","")'
    return (cc[(b + 't', 'Label')], cc[(b + 't', 'Peso')], ntop)


TOPS = {}
TOPS['TA'] = dim_block('TA', 'TipoActivo', 'RankTipoActivo', 8)
TOPS['SE'] = dim_block('SE', 'Sector', 'RankSector', 8)
TOPS['IN'] = dim_block('IN', 'Indicador', 'RankIndicador', 8)
TOPS['TP'] = dim_block('TP', 'TipoPapel', 'RankTipoPapel', 8)
TOPS['EM'] = dim_block('EM', 'NITEmisor', 'RankEmisor', 10, cap=N_EMI, emisor=True)
# nodo (orden estable de maestros)
chdr('NO', ['orden', 'Label', 'Val', 'Peso'])
n_ = {n: cc[('NO', n)] for n in ['orden', 'Label', 'Val', 'Peso']}
for i in range(8):
    r = R0 + i
    ws_c[f"{n_['orden']}{r}"] = i + 1
    ws_c[f"{n_['Label']}{r}"] = f'={q(S_MA)}$J${5 + i}&""'
    ws_c[f"{n_['Val']}{r}"] = f'=IF({n_["Label"]}{r}="",0,SUMIFS({MR("ValorPresente")},{MR("CodigoPortafolio")},{C_SEL},{MR("Nodo")},{n_["Label"]}{r}))'
    ws_c[f"{n_['Peso']}{r}"] = f'=IF({C_TOT}<=0,0,{n_["Val"]}{r}/{C_TOT})'
    ws_c[f"{n_['Peso']}{r}"].number_format = '0.0%'
r = R0 + 8
ws_c[f"{n_['orden']}{r}"] = 9
ws_c[f"{n_['Peso']}{r}"] = f'=IF({C_TOT}<=0,0,MAX(0,ROUND(1-SUM({n_["Peso"]}{R0}:{n_["Peso"]}{r - 1}),6)))'
ws_c[f"{n_['Peso']}{r}"].number_format = '0.0%'
ws_c[f"{n_['Label']}{r}"] = f'=IF({n_["Peso"]}{r}>0.00005,"Fuera de catálogo","")'
TOPS['NO'] = (n_['Label'], n_['Peso'], 9)
ws_c.freeze_panes = 'A5'

# ---- maestros automáticos (dependen del modelo)
for i in range(N_PORT):
    r = 5 + i
    ws_ma.cell(r, 12, i + 1).font = f_norm
    ws_ma.cell(r, 13, f'=IFERROR(INDEX({MR("CodigoPortafolio")},MATCH(L{r},{MR("RankPortafolio")},0)),"")').font = f_norm
    ws_ma.cell(r, 14, f'=IFERROR(INDEX({MR("NombrePortafolio")},MATCH(L{r},{MR("RankPortafolio")},0)),"")').font = f_norm
for i in range(N_EMI):
    r = 5 + i
    ws_ma.cell(r, 23, i + 1).font = f_norm
    for j, n in ((24, 'NITEmisor'), (25, 'Emisor'), (26, 'Sector')):
        ws_ma.cell(r, j, f'=IFERROR(INDEX({MR(n)},MATCH(W{r},{MR("RankEmisor")},0)),"")').font = f_norm
RG_PCOD, RG_PNOM = rg(S_MA, 'M', N_PORT), rg(S_MA, 'N', N_PORT)
wb.defined_names['ListaPortafolios'] = DefinedName(
    'ListaPortafolios', attr_text=f"{q(S_MA)}$M$5:INDEX({q(S_MA)}$M$5:$M${4 + N_PORT},MAX(1,COUNTIF({q(S_MA)}$M$5:$M${4 + N_PORT},\"?*\")))")

# =====================================================================
# 07 RIESGO CALCULADO
# =====================================================================
RCOLS = ['k', 'FilaModelo', 'CodigoPortafolio', 'NITEmisor', 'Emisor', 'ExposicionEmisor', 'ValorTotalPortafolio', 'PesoEmisor',
         'TipoLimite', 'LimitePorcentual', 'LimiteMonto', 'LimiteCalculado', 'Utilizacion', 'Holgura', 'EstadoRiesgo',
         'Calificacion', 'ClaveOrden', 'RankSeleccion']
rc = {n: L(i + 1) for i, n in enumerate(RCOLS)}


def RR(n):
    return rg(S_RC, rc[n], N_PE)


title(ws_rc, 'Riesgo calculado: exposición por portafolio y emisor frente al límite',
      'Una fila por combinación portafolio + emisor. Unidad: la misma de ValorPresente. Filtre por CodigoPortafolio o EstadoRiesgo.')
hdr_row(ws_rc, 4, 1, RCOLS, [5, 9, 14, 13, 30, 15, 15, 9, 20, 10, 13, 14, 10, 13, 24, 11, 10, 10])
ws_rc.row_dimensions[4].height = 30
for i in range(N_PE):
    r = R0 + i
    x = {n: f"{c}{r}" for n, c in rc.items()}
    fm = x['FilaModelo']

    def im(n, alt='""'):
        return f'=IF({fm}=0,{alt},INDEX({MR(n)},{fm}))'
    fr = f'INDEX({MR("FilaRiesgo")},{fm})'
    ws_rc[x['k']] = i + 1
    ws_rc[x['FilaModelo']] = f'=IFERROR(MATCH({x["k"]},{MR("RankPortEmisor")},0),0)'
    ws_rc[x['CodigoPortafolio']] = im('CodigoPortafolio')
    ws_rc[x['NITEmisor']] = im('NITEmisor')
    ws_rc[x['Emisor']] = im('Emisor')
    ws_rc[x['ExposicionEmisor']] = im('ExposicionEmisor')
    ws_rc[x['ValorTotalPortafolio']] = im('ValorTotalPortafolio')
    ws_rc[x['PesoEmisor']] = f'=IF({fm}=0,"",IF({x["ValorTotalPortafolio"]}=0,"",{x["ExposicionEmisor"]}/{x["ValorTotalPortafolio"]}))'
    for n, src in (('TipoLimite', 'TipoLimite'), ('LimitePorcentual', 'Pct'), ('LimiteMonto', 'Monto'), ('Calificacion', 'Calif')):
        ws_rc[x[n]] = f'=IF({fm}=0,"",IF({fr}=0,"",INDEX({CR("R", src, N_RSK)},{fr})))'
    ws_rc[x['LimiteCalculado']] = im('LimiteAplicable')
    ws_rc[x['Utilizacion']] = im('UtilizacionLimite')
    ws_rc[x['Holgura']] = im('HolguraLimite')
    ws_rc[x['EstadoRiesgo']] = im('EstadoRiesgo')
    ws_rc[x['ClaveOrden']] = (f'=IF(OR({fm}=0,{x["CodigoPortafolio"]}<>{C_SEL}),-1,IF({x["EstadoRiesgo"]}="Límite cero",1000,'
                              f'IF(ISNUMBER({x["Utilizacion"]}),10+ROUND({x["Utilizacion"]},6)+IF(ISNUMBER({x["PesoEmisor"]}),{x["PesoEmisor"]},0)*0.0000001,'
                              f'10+{P_U1}-0.000001+IF(ISNUMBER({x["PesoEmisor"]}),MAX(0,MIN(1,{x["PesoEmisor"]})),0)*0.0000001)))')
    ko, kk = RR('ClaveOrden'), RR('k')
    ws_rc[x['RankSeleccion']] = f'=IF({x["ClaveOrden"]}<0,"",SUMPRODUCT(({ko}>{x["ClaveOrden"]})+0)+SUMPRODUCT(({ko}={x["ClaveOrden"]})*({kk}<{x["k"]}))+1)'
    for n, fm_ in (('ExposicionEmisor', NUM2), ('ValorTotalPortafolio', NUM2), ('PesoEmisor', '0.0%'), ('LimitePorcentual', '0.0%'),
                   ('LimiteMonto', NUM2), ('LimiteCalculado', NUM2), ('Utilizacion', '0.0%'), ('Holgura', NUM2)):
        ws_rc[x[n]].number_format = fm_
add_table(ws_rc, 'tblRiesgoCalculado', f"A4:{L(len(RCOLS))}{R0 + N_PE - 1}", 'TableStyleLight9')
ws_rc.freeze_panes = 'F5'

# =====================================================================
# 06 ACTIVOS COMUNES
# =====================================================================
ACOLS = ['k', 'FilaModelo', 'ClaveActivo', 'Nemotecnico', 'Emisor', 'TipoActivo', 'Indicador', 'Nodo', 'NumPortafoliosTotal',
         'NumPortafoliosSeleccionados', 'ValorSeleccionados', 'NumAumentar', 'NumReducir', 'NumMantener', 'NumSinDefinir',
         'Consenso', 'ListaPortafolios', 'ClaveOrden', 'Rank']
ac = {n: L(i + 1) for i, n in enumerate(ACOLS)}


def AR(n):
    return rg(S_AC, ac[n], N_ACT)


title(ws_ac, 'Activos comunes entre los portafolios seleccionados en el dashboard',
      'Una fila por activo (ClaveActivo). Un activo es común si lo tienen 2 o más de los portafolios seleccionados. Filtre por Consenso.')
hdr_row(ws_ac, 4, 1, ACOLS, [5, 9, 26, 16, 28, 22, 11, 22, 12, 14, 15, 10, 10, 10, 10, 24, 34, 10, 7])
ws_ac.row_dimensions[4].height = 40
for i in range(N_ACT):
    r = R0 + i
    x = {n: f"{c}{r}" for n, c in ac.items()}
    fm, cl = x['FilaModelo'], x['ClaveActivo']
    ws_ac[x['k']] = i + 1
    ws_ac[fm] = f'=IFERROR(MATCH({x["k"]},{MR("RankActivo")},0),0)'
    for n in ('ClaveActivo', 'Nemotecnico', 'Emisor', 'TipoActivo', 'Indicador', 'Nodo'):
        ws_ac[x[n]] = f'=IF({fm}=0,"",INDEX({MR(n)},{fm}))'

    def sm(col):
        return f'=IF({fm}=0,"",SUMIFS({MR(col)},{MR("ClaveActivo")},{cl}))'
    ws_ac[x['NumPortafoliosTotal']] = sm('PrimPA')
    ws_ac[x['NumPortafoliosSeleccionados']] = sm('SelPA')
    ws_ac[x['ValorSeleccionados']] = f'=IF({fm}=0,"",SUMIFS({MR("ValorPresente")},{MR("ClaveActivo")},{cl},{MR("Seleccionado")},1))'
    ws_ac[x['ValorSeleccionados']].number_format = NUM2
    ws_ac[x['NumAumentar']] = sm('SelAumentar')
    ws_ac[x['NumReducir']] = sm('SelReducir')
    ws_ac[x['NumMantener']] = sm('SelMantener')
    ws_ac[x['NumSinDefinir']] = sm('SelSinDefinir')
    J, A, R_, M_, S_ = x['NumPortafoliosSeleccionados'], x['NumAumentar'], x['NumReducir'], x['NumMantener'], x['NumSinDefinir']
    ws_ac[x['Consenso']] = (f'=IF({fm}=0,"",IF({J}<2,"No común en la selección",IF(AND({A}>0,{R_}>0),"Conflicto",'
                            f'IF({S_}>0,"Información incompleta",IF({A}={J},"Compra conjunta",IF({R_}={J},"Venta conjunta",'
                            f'IF({M_}={J},"Mantener",IF({A}>0,"Sesgo comprador","Sesgo vendedor"))))))))')
    parts = '&'.join(
        f'IF({cell(S_D, s, 11)}="","",IF(COUNTIF({MR("ClavePortafolioActivo")},{cell(S_D, s, 11)}&"|"&{cl})>0,{cell(S_D, s, 11)}&"; ",""))'
        for s in SLOTS)
    ws_ac[x['ListaPortafolios']] = f'=IF({fm}=0,"",IF({J}<2,"",{parts}))'
    ws_ac[x['ClaveOrden']] = f'=IF({fm}=0,-1,IF({J}>=2,{J}+MAX(0,MIN(0.999999,{x["ValorSeleccionados"]}/({C_TOTSEL}+1))),-1))'
    ko, kk = AR('ClaveOrden'), AR('k')
    ws_ac[x['Rank']] = f'=IF({x["ClaveOrden"]}<0,"",SUMPRODUCT(({ko}>{x["ClaveOrden"]})+0)+SUMPRODUCT(({ko}={x["ClaveOrden"]})*({kk}<{x["k"]}))+1)'
add_table(ws_ac, 'tblActivosComunes', f"A4:{L(len(ACOLS))}{R0 + N_ACT - 1}", 'TableStyleLight9')
ws_ac.freeze_panes = 'E5'

# =====================================================================
# 08 CALIDAD
# =====================================================================
title(ws_q, 'Calidad de datos: controles y excepciones',
      'Arriba: resumen por control. Abajo: detalle fila a fila (FilaOrigen = fila en 01_POSICIONES_INPUT).')
hdr_row(ws_q, 4, 1, ['Control', 'Severidad', 'Cantidad', 'Estado', 'DondeRevisar'], [58, 15, 10, 12, 60])
qrows = [(lab, sev, f'=SUM({MR(n)})', 'Detalle en esta hoja (tabla inferior) y columna %s de 05_MODELO_POSICIONES' % n) for n, lab, sev in QFLAGS]
P_ = q(S_P)
qrows += [
    ('Fecha de posiciones diferente de fecha de riesgo', 'Advertencia',
     f'=IF(OR(COUNT({MR("FechaCorte")})=0,COUNT({q(S_R)}$A:$A)=0),0,IF(MAX({MR("FechaCorte")})<>MAX({q(S_R)}$A:$A),1,0))',
     'Compare FechaCorte en 01_POSICIONES_INPUT y 03_RIESGO_INPUT'),
    ('Fecha de posiciones diferente de fecha de estrategia', 'Advertencia',
     f'=IF(OR(COUNT({MR("FechaCorte")})=0,COUNT({q(S_E)}$A:$A)=0),0,IF(MAX({MR("FechaCorte")})<>MAX({q(S_E)}$A:$A),1,0))',
     'Compare FechaCorte en 01_POSICIONES_INPUT y 02_ESTRATEGIA_INPUT'),
    ('Más de una FechaCorte en la sábana de posiciones', 'Advertencia',
     f'=IF(COUNT({MR("FechaCorte")})=0,0,IF(MAX({MR("FechaCorte")})<>MIN({MR("FechaCorte")}),1,0))', '01_POSICIONES_INPUT, columna FechaCorte'),
    ('Claves de estrategia duplicadas (portafolio + dimensión + categoría)', 'Error',
     f'=SUM({CR("S", "Dup", N_STR)})', '02_ESTRATEGIA_INPUT: se usa la primera coincidencia'),
    ('Estrategia con DimensionEstrategia no reconocida', 'Advertencia',
     f'=SUM({CR("S", "DimInv", N_STR)})', '02_ESTRATEGIA_INPUT: dimensiones válidas en 04_MAESTROS!E5:E11'),
    ('Claves de riesgo duplicadas (portafolio + NIT)', 'Error',
     f'=SUM({CR("R", "Dup", N_RSK)})', '03_RIESGO_INPUT: se usa la primera coincidencia'),
    ('Capacidad del modelo excedida (filas que no se están leyendo)', 'Error crítico',
     (f'=IF(COUNTA({P_}$B:$B)-1>{N_POS},1,0)+IF(COUNTA({q(S_E)}$B:$B)-1>{N_STR},1,0)+IF(COUNTA({q(S_R)}$B:$B)-1>{N_RSK},1,0)'
      f'+IF(MAX({MR("RankActivo")})>{N_ACT},1,0)+IF(MAX({MR("RankPortEmisor")})>{N_PE},1,0)+IF(MAX({MR("RankEmisor")})>{N_EMI},1,0)'
      f'+IF(MAX({MR("RankPortafolio")})>{N_PORT},1,0)'),
     f'Capacidades: posiciones {N_POS}, estrategia {N_STR}, riesgo {N_RSK}, activos {N_ACT}, portafolio-emisor {N_PE}, emisores {N_EMI}, portafolios {N_PORT}'),
]
for i, (lab, sev, fml, donde) in enumerate(qrows, 5):
    ws_q.cell(i, 1, lab).font = f_norm
    ws_q.cell(i, 2, sev).font = f_norm
    ws_q.cell(i, 3, fml).font = f_bold
    ws_q.cell(i, 4, f'=IF(C{i}=0,"OK","Revisar")').font = f_norm
    ws_q.cell(i, 5, donde).font = Font(name=FN, size=8, color='595959')
QR1, QR2 = 5, 4 + len(qrows)
add_table(ws_q, 'tblCalidadDatos', f"A4:E{QR2}", 'TableStyleLight9')
RG_QSEV, RG_QCNT = f"{q(S_Q)}$B${QR1}:$B${QR2}", f"{q(S_Q)}$C${QR1}:$C${QR2}"
QD0 = QR2 + 5
ws_q.cell(QD0 - 2, 1, 'Detalle de filas con incidencia').font = Font(name=FN, size=11, bold=True, color=NAVY)
ws_q.cell(QD0 - 2, 3, f'=MAX({MR("RankIncidencia")})').font = f_bold
ws_q.cell(QD0 - 2, 4, f'filas con incidencia (se listan hasta {N_QD})').font = f_norm
hdr_row(ws_q, QD0 - 1, 1, ['Detalle', 'Severidad', 'FilaOrigen', 'CodigoPortafolio', 'Nemotecnico / ClaveActivo', 'k'])
ws_q.column_dimensions['F'].width = 6
for i in range(N_QD):
    r = QD0 + i
    fm = f'IFERROR(MATCH(F{r},{MR("RankIncidencia")},0),0)'
    ws_q.cell(r, 6, i + 1)
    ws_q.cell(r, 1, f'=IF({fm}=0,"",INDEX({MR("DetalleCalidad")},{fm}))')
    ws_q.cell(r, 2, f'=IF({fm}=0,"",INDEX({MR("EstadoCalidad")},{fm}))')
    ws_q.cell(r, 3, f'=IF({fm}=0,"",INDEX({MR("FilaOrigen")},{fm}))')
    ws_q.cell(r, 4, f'=IF({fm}=0,"",TRIM(INDEX({P_}$B:$B,INDEX({MR("FilaOrigen")},{fm}))&""))')
    ws_q.cell(r, 5, f'=IF({fm}=0,"",TRIM(INDEX({P_}$F:$F,INDEX({MR("FilaOrigen")},{fm}))&"")&" / "&INDEX({MR("ClaveActivo")},{fm}))')
    for j in range(1, 7):
        ws_q.cell(r, j).font = f_norm
add_table(ws_q, 'tblCalidadDetalle', f"A{QD0 - 1}:F{QD0 + N_QD - 1}", 'TableStyleLight9')
ws_q.freeze_panes = 'A5'

exec(open(__file__.replace('build.py', 'dash.py'), encoding='utf-8').read())

if os.environ.get('HIDE'):
    for w in wb.worksheets[1:]:
        w.sheet_state = 'hidden'
wb.active = 0
for w in wb.worksheets:
    w.sheet_view.tabSelected = (w.title == S_D)
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
import zipfile, shutil
tmp = OUT + '.tmp'
with zipfile.ZipFile(OUT) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        data = zi.read(it.filename)
        if it.filename.startswith('xl/charts/chart'):
            data = data.replace(b'<numFmt formatCode="0.0%;;"/>', b'<numFmt formatCode="0.0%;;" sourceLinked="0"/>')
        zo.writestr(it, data)
shutil.move(tmp, OUT)
print('saved', OUT)
