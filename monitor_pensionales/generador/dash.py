# ---- se ejecuta dentro del espacio de nombres de build.py ----
D = ws_d
D.sheet_view.showGridLines = False
D.sheet_properties.tabColor = NAVY
D.column_dimensions['A'].width = 2
for j in range(2, 26):
    D.column_dimensions[L(j)].width = 9
D.column_dimensions['Z'].width = 2
D.column_dimensions['AA'].hidden = True
D.column_dimensions['AB'].hidden = True
D.sheet_view.zoomScale = 85
white = Font(name=FN, size=9, color='FFFFFF')
grey_s = Font(name=FN, size=8, italic=True, color='595959')
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=False, indent=0)
RIGHT = Alignment(horizontal='right', vertical='center')


def put(rng_, value=None, font=f_norm, fill=None, align=LEFT, fmt=None, border=None):
    a = rng_.split(':')[0]
    if ':' in rng_:
        D.merge_cells(rng_)
        for row in D[rng_]:
            for c in row:
                if fill: c.fill = fill
                if border: c.border = border
    c = D[a]
    if value is not None: c.value = value
    c.font = font
    if fill: c.fill = fill
    c.alignment = align
    if fmt: c.number_format = fmt
    if border: c.border = border
    return c


def bar(r, text, c1='B', c2='Y'):
    put(f"{c1}{r}:{c2}{r}", text, Font(name=FN, size=10, bold=True, color='FFFFFF'), fill_teal, LEFT)


def colhdr(r, spans):
    for rng_, text in spans:
        put(f"{rng_.split(':')[0]}{r}:{rng_.split(':')[-1]}{r}" if ':' in rng_ else f"{rng_}{r}", text, f_hdr, fill_hdr, CEN)
    D.row_dimensions[r].height = 26


def helper(r, formula, col='AA'):
    c = D[f"{col}{r}"]; c.value = formula; c.font = Font(name=FN, size=7, color='BFBFBF')
    return f"${col}{r}"


SEL = C_SEL
VP, COD = MR('ValorPresente'), MR('CodigoPortafolio')
ALERTAS = ('Alerta', 'Al límite o excedido', 'Límite cero')
SINCRUCE = ('Sin información de riesgo', 'Sin límite informado')
ws_c['D1'] = f"=IF({q(S_D)}$B$8=\"\",\"(ninguno)\",{q(S_D)}$B$8&\"\")"

# ---------------- encabezado ----------------
for r in (2, 3):
    for j in range(2, 26):
        D.cell(r, j).fill = fill_hdr
    D.row_dimensions[r].height = 20
put('B2:M3', 'MONITOR PENSIONALES', Font(name=FN, size=20, bold=True, color='FFFFFF'), fill_hdr, LEFT)
put('N2:O2', 'Fecha de corte', white, fill_hdr, CEN)
put('N3:O3', f'=IF(COUNT({MR("FechaCorte")})=0,"Sin dato",MAX({MR("FechaCorte")}))', Font(name=FN, size=11, bold=True, color='FFFFFF'), fill_hdr, CEN, DATE)
put('P2:Q2', 'Fecha de riesgo', white, fill_hdr, CEN)
put('P3:Q3', f'=IF(COUNT({q(S_R)}$A:$A)=0,"Sin dato",MAX({q(S_R)}$A:$A))', Font(name=FN, size=11, bold=True, color='FFFFFF'), fill_hdr, CEN, DATE)
put('R2:Y2', 'Estado de actualización', white, fill_hdr, CEN)
_err = f'(SUMIFS({RG_QCNT},{RG_QSEV},"Error")+SUMIFS({RG_QCNT},{RG_QSEV},"Error crítico"))'
_adv = f'SUMIFS({RG_QCNT},{RG_QSEV},"Advertencia")'
put('R3:Y3', (f'=IF(COUNT({MR("FechaCorte")})=0,"Sin posiciones cargadas",IF({_err}>0,"Revisar. Errores: "&{_err}&" | Advertencias: "&{_adv}&" (hoja 08_CALIDAD_DATOS)",'
              f'IF({_adv}>0,"Actualizado. Advertencias: "&{_adv}&" (hoja 08_CALIDAD_DATOS)","Actualizado sin incidencias")))'),
    Font(name=FN, size=9, bold=True, color='FFFFFF'), fill_hdr, CEN)
put('B4:Y4', f'=IF({P_MODO}="Sintético","Datos sintéticos para demostración. No usar para decisiones de inversión.","")',
    Font(name=FN, size=10, bold=True, color='C00000'), None, CEN)
D.conditional_formatting.add('B4:Y4', FormulaRule(formula=['$B$4<>""'], fill=PatternFill('solid', bgColor='FFE699', fgColor='FFE699')))

# ---------------- filtros ----------------
bar(6, 'FILTROS  (celdas amarillas = seleccionables)')
lab = Font(name=FN, size=8, bold=True, color=NAVY)
put('B7:D7', 'Portafolio principal', lab)
put('B8:D8', os.environ.get('SEL', 'PEN001'), Font(name=FN, size=11, bold=True, color='0000FF'), fill_in, CEN, border=b_all)
put('F7:N7', 'Nombre completo', lab)
put('F8:N8', f'=IF($B$8="","Seleccione un portafolio",IFERROR(INDEX({RG_PNOM},MATCH($B$8&"",{RG_PCOD},0)),"Portafolio no encontrado en la sábana"))',
    Font(name=FN, size=11, bold=True, color=NAVY), fill_lb, LEFT)
put('P7:R7', 'Dimensión de estrategia', lab)
put('P8:R8', os.environ.get('DIMSEL', 'Activo'), Font(name=FN, size=11, bold=True, color='0000FF'), fill_in, CEN, border=b_all)
put('T7:Y7', 'Actualización', lab)
put('T8:Y8', 'Sin macros: pegue las entradas y se recalcula solo (F9 si el cálculo está en manual).', grey_s, None,
    Alignment(horizontal='left', vertical='center', wrap_text=True))
D.row_dimensions[8].height = 24
put('B9:Y9', 'Portafolios a comparar (hasta 6). Deje vacío lo que no use. Con ellos se calculan los activos comunes y el consenso.', grey_s)
defaults = [x or None for x in os.environ.get('SLOTSEL', 'PEN001,PEN002,PEN003,PEN004,,').split(',')]
for k, s in enumerate(SLOTS):
    e = L(ord(s) - 64 + 3)
    put(f"{s}10:{e}10", f'Comparar {k + 1}', lab, None, CEN)
    put(f"{s}11:{e}11", defaults[k], Font(name=FN, size=10, bold=True, color='0000FF'), fill_in, CEN, border=b_all)
dvp = DataValidation(type='list', formula1='=ListaPortafolios', allow_blank=True)
dvp.error, dvp.errorTitle = 'Elija un código de la lista de portafolios Pensionales.', 'Portafolio no válido'
D.add_data_validation(dvp)
dvp.add('B8')
for s in SLOTS:
    dvp.add(f"{s}11")
dvd = DataValidation(type='list', formula1=f"={q(S_MA)}$E$5:$E$10", allow_blank=False)
D.add_data_validation(dvd); dvd.add('P8')

# ---------------- KPI ----------------
bar(13, 'INDICADORES DEL PORTAFOLIO PRINCIPAL')
RCOD, REST, REXP = RR('CodigoPortafolio'), RR('EstadoRiesgo'), RR('ExposicionEmisor')
SCOD, SDIM, SDIR = CR('S', 'Cod', N_STR), CR('S', 'Dim', N_STR), CR('S', 'Dir', N_STR)


def sum_estados(code, estados, what='sum'):
    if what == 'sum':
        return '(' + '+'.join(f'SUMIFS({REXP},{RCOD},{code},{REST},"{e}")' for e in estados) + ')'
    return '(' + '+'.join(f'COUNTIFS({RCOD},{code},{REST},"{e}")' for e in estados) + ')'


nsel = f'SUMPRODUCT(({RG_PCOD}<>"")*(COUNTIF({RG_SLOTS},{RG_PCOD})>0))'
kpis = [
    ('Valor de mercado (COP MM)', f'={C_TOT}*{P_FAC}', NUM, f'=COUNTIF({COD},{SEL})&" posiciones"'),
    ('Emisores', f'=COUNTIF({RCOD},{SEL})', '0', 'con exposición'),
    ('Activos distintos', f'=SUMIFS({MR("PrimPA")},{COD},{SEL})', '0', 'por ClaveActivo'),
    ('% exposición en alerta o excedida', f'=IF({C_TOT}<=0,0,{sum_estados(SEL, ALERTAS)}/{C_TOT})', '0.0%', f'={sum_estados(SEL, ALERTAS, "n")}&" emisores"'),
    ('Aumentar / Reducir', f'=COUNTIFS({SCOD},{SEL},{SDIM},{P_DIMN},{SDIR},"Aumentar")&" / "&COUNTIFS({SCOD},{SEL},{SDIM},{P_DIMN},{SDIR},"Reducir")',
     None, f'="categorías ("&{P_DIMN}&")"'),
    ('Emisores sin cruce de riesgo', f'={sum_estados(SEL, SINCRUCE, "n")}', '0', 'requieren revisión manual'),
    ('Activos comunes', f'=COUNTIF({AR("NumPortafoliosSeleccionados")},">=2")', '0', f'="entre "&{nsel}&" portafolios comparados"'),
    ('Conflictos de estrategia', f'=COUNTIF({AR("Consenso")},"Conflicto")', '0', 'en activos comunes'),
]
for k, (lb, fml, fmt, sub) in enumerate(kpis):
    c1, c2 = L(2 + 3 * k), L(4 + 3 * k)
    put(f"{c1}14:{c2}14", lb, Font(name=FN, size=8, bold=True, color=NAVY), fill_lb, CEN, border=b_all)
    put(f"{c1}15:{c2}15", fml, Font(name=FN, size=16, bold=True, color=NAVY), None, CEN, fmt, border=b_all)
    put(f"{c1}16:{c2}16", sub, grey_s, None, CEN, border=b_all)
D.row_dimensions[14].height = 24
D.row_dimensions[15].height = 30
for c_ in ('K15', 'Q15', 'W15'):   # alertas: rojo si > 0
    D.conditional_formatting.add(c_, CellIsRule(operator='greaterThan', formula=['0'], font=Font(name=FN, bold=True, color='C00000')))

# ---------------- gráficos ----------------
bar(18, 'COMPOSICIÓN DEL PORTAFOLIO PRINCIPAL  (% del valor de mercado)')


def chart(block, anchor_col, row, title_):
    lc, pc, n = TOPS[block]
    lci, pci = [__import__('openpyxl').utils.column_index_from_string(x) for x in (lc, pc)]
    ch = BarChart()
    ch.type, ch.style, ch.legend = 'bar', 10, None
    ch.add_data(Reference(ws_c, min_col=pci, min_row=R0, max_row=R0 + n - 1), titles_from_data=False)
    ch.set_categories(Reference(ws_c, min_col=lci, min_row=R0, max_row=R0 + n - 1))
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill=TEAL)
    s.graphicalProperties.line.solidFill = TEAL
    ch.dataLabels = DataLabelList(numFmt='0.0%;;', showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
    ch.gapWidth = 45
    ch.x_axis.scaling.orientation = 'maxMin'
    ch.x_axis.delete = False
    ch.y_axis.delete = True
    ch.y_axis.majorGridlines = None
    ch.y_axis.scaling.min = 0
    ch.x_axis.number_format = '@'
    ch.width, ch.height = 14.6, 6.7
    ch.graphical_properties = GraphicalProperties(noFill=True)
    ch.graphical_properties.line.noFill = True
    D.add_chart(ch, f"{anchor_col}{row + 1}")
    e = L(__import__('openpyxl').utils.column_index_from_string(anchor_col) + 7)
    put(f"{anchor_col}{row}:{e}{row}", title_, Font(name=FN, size=9, bold=True, color=NAVY), fill_lb, CEN)


chart('TA', 'B', 19, 'Por activo'); chart('SE', 'J', 19, 'Por sector'); chart('EM', 'R', 19, 'Por emisor (principales y Otros)')
chart('NO', 'B', 33, 'Por nodo (orden de plazo)'); chart('IN', 'J', 33, 'Por indicador'); chart('TP', 'R', 33, 'Por tipo de papel')

# ---------------- riesgo ----------------
RT = 48
bar(RT, 'RIESGO: EXPOSICIÓN POR EMISOR FRENTE AL LÍMITE  (COP MM)', 'B', 'L')
colhdr(RT + 1, [('B:D', 'Emisor'), ('E:F', 'Exposición'), ('G:H', 'Límite'), ('I', 'Utilización'), ('J', 'Holgura'), ('K:L', 'Estado')])
for i in range(1, 9):
    r = RT + 1 + i
    ix = helper(r, f'=IFERROR(MATCH({i},{RR("RankSeleccion")},0),0)')
    msg = '"Sin emisores para el portafolio seleccionado"' if i == 1 else '""'
    put(f"B{r}:D{r}", f'=IF({ix}=0,{msg},INDEX({RR("Emisor")},{ix}))', border=b_bot)
    put(f"E{r}:F{r}", f'=IF({ix}=0,"",INDEX({REXP},{ix})*{P_FAC})', align=RIGHT, fmt=NUM, border=b_bot)
    put(f"G{r}:H{r}", f'=IF({ix}=0,"",IF(ISNUMBER(INDEX({RR("LimiteCalculado")},{ix})),INDEX({RR("LimiteCalculado")},{ix})*{P_FAC},"Sin límite"))',
        align=RIGHT, fmt='#,##0;(#,##0);0', border=b_bot)
    put(f"I{r}", f'=IF({ix}=0,"",IF(ISNUMBER(INDEX({RR("Utilizacion")},{ix})),INDEX({RR("Utilizacion")},{ix}),"n/d"))', align=RIGHT, fmt='0.0%', border=b_bot)
    put(f"J{r}", f'=IF({ix}=0,"",IF(ISNUMBER(INDEX({RR("Holgura")},{ix})),INDEX({RR("Holgura")},{ix})*{P_FAC},"n/d"))', align=RIGHT,
        fmt='#,##0;[Red](#,##0);"-"', border=b_bot)
    put(f"K{r}:L{r}", f'=IF({ix}=0,"",INDEX({REST},{ix}))', Font(name=FN, size=8, bold=True), None, CEN, border=b_bot)
r = RT + 10
put(f"B{r}:L{r}", f'="Se muestran 8 de "&COUNTIF({RCOD},{SEL})&" emisores: primero mayor utilización; los que no tienen cruce de riesgo van antes de los que están dentro de límite. Lista completa en 07_RIESGO_CALCULADO."', Font(name=FN, size=7, italic=True, color='595959'), None, Alignment(wrap_text=True, vertical='top'))
D.row_dimensions[r].height = 22

# ---------------- estrategia ----------------
bar(RT, 'ESTRATEGIA DEL PORTAFOLIO PRINCIPAL  (dimensión elegida y duración)', 'N', 'Y')
colhdr(RT + 1, [('N', 'Dimensión'), ('O:R', 'Categoría'), ('S', 'Peso actual'), ('T', 'Estrategia'), ('U:V', 'Unidad'), ('W', 'Dirección'), ('X:Y', 'Peso objetivo')])
dimcols = [('Activo', 'TipoActivo'), ('Nodo', 'Nodo'), ('Indicador', 'Indicador'), ('TipoPapel', 'TipoPapel'), ('Sector', 'Sector'), ('Emisor', 'Emisor')]
for i in range(1, 9):
    r = RT + 1 + i
    ix = helper(r, f'=IFERROR(MATCH({i},{CR("S", "Rank", N_STR)},0),0)', 'AB')
    msg = '"Sin estrategia cargada para el portafolio"' if i == 1 else '""'
    put(f"N{r}", f'=IF({ix}=0,"",INDEX({SDIM},{ix}))', border=b_bot)
    put(f"O{r}:R{r}", f'=IF({ix}=0,{msg},INDEX({CR("S", "Cat", N_STR)},{ix}))', border=b_bot)
    nest = '"n/d"'
    for dn, col in reversed(dimcols):
        nest = f'IF($N{r}="{dn}",SUMIFS({VP},{COD},{SEL},{MR(col)},$O{r})/{C_TOT},{nest})'
    put(f"S{r}", f'=IF(OR({ix}=0,{C_TOT}<=0),"",{nest})', align=RIGHT, fmt='0.0%', border=b_bot)
    put(f"T{r}", f'=IF({ix}=0,"",INDEX({CR("S", "Val", N_STR)},{ix}))', align=RIGHT, fmt='+0.00;[Red]-0.00;0.00', border=b_bot)
    put(f"U{r}:V{r}", f'=IF({ix}=0,"",INDEX({CR("S", "Unidad", N_STR)},{ix}))', Font(name=FN, size=8), None, CEN, border=b_bot)
    dr = f'INDEX({SDIR},{ix})'
    put(f"W{r}", f'=IF({ix}=0,"",IF($N{r}="Duracion",IF({dr}="Aumentar","Alargar",IF({dr}="Reducir","Acortar",{dr})),{dr}))',
        Font(name=FN, size=8, bold=True), None, CEN, border=b_bot)
    put(f"X{r}:Y{r}", (f'=IF(OR(NOT(ISNUMBER($S{r})),NOT(ISNUMBER($T{r}))),"",IF($U{r}="PuntosPorcentuales",$S{r}+$T{r}/100,'
                       f'IF($U{r}="COP_MM",$S{r}+$T{r}/({C_TOT}*{P_FAC}),"")))'), align=RIGHT, fmt='0.0%', border=b_bot)
r = RT + 10
put(f"N{r}:V{r}", f'="Peso del portafolio sin estrategia definida (dimensión "&{P_DIMN}&"):"', grey_s)
put(f"W{r}", f'=IF({C_TOT}<=0,0,SUMIFS({VP},{COD},{SEL},{MR("DireccionEstrategia")},"Sin definir")/{C_TOT})', Font(name=FN, size=9, bold=True), None, RIGHT, '0.0%')
put(f"X{r}:Y{r}", f'=COUNTIF({SCOD},{SEL})&" líneas en total"', grey_s, None, RIGHT)

# ---------------- comparación ----------------
CT = 60
bar(CT, 'COMPARACIÓN DE LOS PORTAFOLIOS SELECCIONADOS: ¿ALARGAR, ACORTAR O MANTENER?')
colhdr(CT + 1, [('B:C', 'Código'), ('D:I', 'Nombre'), ('J:K', 'Valor mercado (COP MM)'), ('L', '# activos'), ('M:N', 'Duración actual'),
                ('O:P', 'Estrategia duración'), ('Q:R', 'Dirección duración'), ('S', '# Aumentar'), ('T', '# Reducir'),
                ('U', '# Mantener'), ('V', '# Sin definir'), ('W:X', 'Activos comunes'), ('Y', '% en alerta')])
for k, s in enumerate(SLOTS):
    r = CT + 2 + k
    c = cell(S_D, s, 11)
    ix = helper(r, f'=IF({c}="",0,IFERROR(MATCH({c}&"|Duracion|*",{CR("S", "Key", N_STR)},0),0))')
    tot = f'SUMIFS({VP},{COD},{c})'
    put(f"B{r}:C{r}", f'=IF({c}="","",{c}&"")', f_bold, None, CEN, border=b_bot)
    put(f"D{r}:I{r}", f'=IF({c}="","",IFERROR(INDEX({RG_PNOM},MATCH({c}&"",{RG_PCOD},0)),"No encontrado en la sábana"))', border=b_bot)
    put(f"J{r}:K{r}", f'=IF({c}="","",{tot}*{P_FAC})', align=RIGHT, fmt=NUM, border=b_bot)
    put(f"L{r}", f'=IF({c}="","",SUMIFS({MR("PrimPA")},{COD},{c}))', align=CEN, fmt='0', border=b_bot)
    put(f"M{r}:N{r}", f'=IF({c}="","",IF(SUMIFS({MR("VPconDur")},{COD},{c})=0,"n/d",SUMIFS({MR("VPxDur")},{COD},{c})/SUMIFS({MR("VPconDur")},{COD},{c})))',
        align=CEN, fmt='0.00', border=b_bot)
    put(f"O{r}:P{r}", f'=IF(OR({c}="",{ix}=0),"",INDEX({CR("S", "Val", N_STR)},{ix}))', align=CEN, fmt='+0.00;[Red]-0.00;0.00', border=b_bot)
    dr = f'INDEX({SDIR},{ix})'
    put(f"Q{r}:R{r}", f'=IF({c}="","",IF({ix}=0,"Sin definir",IF({dr}="Aumentar","Alargar",IF({dr}="Reducir","Acortar",{dr}))))',
        Font(name=FN, size=8, bold=True), None, CEN, border=b_bot)
    for col, w in (('S', 'Aumentar'), ('T', 'Reducir'), ('U', 'Mantener'), ('V', 'Sin definir')):
        put(f"{col}{r}", f'=IF({c}="","",COUNTIFS({SCOD},{c},{SDIM},{P_DIMN},{SDIR},"{w}"))', align=CEN, fmt='0', border=b_bot)
    put(f"W{r}:X{r}", f'=IF({c}="","",COUNTIFS({COD},{c},{MR("NumSelActivo")},">=2"))', align=CEN, fmt='0', border=b_bot)
    put(f"Y{r}", f'=IF({c}="","",IF({tot}<=0,0,{sum_estados(c, ALERTAS)}/{tot}))', align=CEN, fmt='0.0%', border=b_bot)
r = CT + 8
q1, q2 = f"$Q${CT + 2}:$R${CT + 7}", None
nA, nR, nM, nS = (f'COUNTIF({q1},"{w}")' for w in ('Alargar', 'Acortar', 'Mantener', 'Sin definir'))
nn = f'({nA}+{nR}+{nM}+{nS})'
put(f"B{r}:E{r}", 'Consenso de duración:', f_bold)
put(f"F{r}:L{r}", (f'=IF({nn}=0,"Sin portafolios seleccionados",IF(AND({nA}>0,{nR}>0),"Conflicto: unos alargan y otros acortan",'
                   f'IF({nS}>0,"Información incompleta",IF({nA}={nn},"Todos alargan",IF({nR}={nn},"Todos acortan",'
                   f'IF({nM}={nn},"Todos mantienen",IF({nA}>0,"Sesgo a alargar","Sesgo a acortar")))))))'), Font(name=FN, size=10, bold=True), None, LEFT)
put(f"N{r}:Y{r}", f'="Los conteos # Aumentar, # Reducir, # Mantener y # Sin definir son líneas de estrategia de la dimensión "&{P_DIMN}&"."', grey_s)

# ---------------- activos comunes ----------------
AT = 70
bar(AT, 'ACTIVOS COMUNES ENTRE LOS PORTAFOLIOS SELECCIONADOS Y CONSENSO DE ESTRATEGIA')
pairs = [('N', 'O'), ('P', 'Q'), ('R', 'S'), ('T', 'U'), ('V', 'W'), ('X', 'Y')]
colhdr(AT + 1, [('B:C', 'Nemotécnico'), ('D:F', 'Emisor'), ('G:I', 'ClaveActivo'), ('J', '# Port.'), ('K:M', 'Consenso')])
for (a, b), s in zip(pairs, SLOTS):
    put(f"{a}{AT + 1}:{b}{AT + 1}", f'=IF({cell(S_D, s, 11)}="","(vacío)",{cell(S_D, s, 11)}&" (dirección y peso)")', f_hdr, fill_hdr, CEN)
NCOM = 12
for i in range(1, NCOM + 1):
    r = AT + 1 + i
    ix = helper(r, f'=IFERROR(MATCH({i},{AR("Rank")},0),0)')
    msg = '"No hay activos comunes entre los portafolios seleccionados"' if i == 1 else '""'
    put(f"B{r}:C{r}", f'=IF({ix}=0,"",INDEX({AR("Nemotecnico")},{ix}))', f_bold, border=b_bot)
    put(f"D{r}:F{r}", f'=IF({ix}=0,{msg},INDEX({AR("Emisor")},{ix}))', border=b_bot)
    put(f"G{r}:I{r}", f'=IF({ix}=0,"",INDEX({AR("ClaveActivo")},{ix}))', Font(name=FN, size=8), border=b_bot)
    put(f"J{r}", f'=IF({ix}=0,"",INDEX({AR("NumPortafoliosSeleccionados")},{ix}))', align=CEN, fmt='0', border=b_bot)
    put(f"K{r}:M{r}", f'=IF({ix}=0,"",INDEX({AR("Consenso")},{ix}))', Font(name=FN, size=8, bold=True), None, CEN, border=b_bot)
    for (a, b), s in zip(pairs, SLOTS):
        c = cell(S_D, s, 11)
        key = f'{c}&"|"&$G{r}'
        put(f"{a}{r}", f'=IF(OR({ix}=0,{c}=""),"",IFERROR(INDEX({MR("DireccionEstrategia")},MATCH({key},{MR("ClavePortafolioActivo")},0)),"-"))',
            Font(name=FN, size=8, bold=True), None, CEN, border=b_bot)
        put(f"{b}{r}", (f'=IF(OR({a}{r}="",{a}{r}="-"),"",IF(SUMIFS({VP},{COD},{c})=0,"",'
                        f'SUMIFS({VP},{MR("ClavePortafolioActivo")},{key})/SUMIFS({VP},{COD},{c})))'), align=CEN, fmt='0.0%', border=b_bot)
r = AT + NCOM + 2
cons = ['Compra conjunta', 'Venta conjunta', 'Mantener', 'Sesgo comprador', 'Sesgo vendedor', 'Conflicto', 'Información incompleta']
for k, w in enumerate(cons):
    c1, c2, c3 = L(2 + 3 * k), L(3 + 3 * k), L(4 + 3 * k)
    put(f"{c1}{r}:{c2}{r}", w, Font(name=FN, size=8, bold=True), None, CEN, border=b_all)
    put(f"{c3}{r}", f'=COUNTIF({AR("Consenso")},"{w}")', f_bold, None, CEN, '0', border=b_all)
put(f"W{r}:Y{r}", '= activos por consenso', grey_s)
D[f"W{r}"].value = 'activos por tipo de consenso'
r += 1
put(f"B{r}:Y{r}", (f'="Se muestran hasta {NCOM} de "&COUNTIF({AR("NumPortafoliosSeleccionados")},">=2")&" activos comunes, ordenados por número de portafolios y valor. '
                   f'Lista completa con filtros en 06_ACTIVOS_COMUNES. Un guion (-) indica que ese portafolio no tiene el activo."'), grey_s)
LAST = r + 2
put(f"B{LAST}:Y{LAST}", ('Aumentar / Alargar = estrategia positiva.  Reducir / Acortar = estrategia negativa.  Mantener = cero explícito.  '
                         'Sin definir = no hay instrucción.  La dirección de cada activo se toma de la categoría de la dimensión de estrategia elegida.'), grey_s)

# ---------------- formato condicional ----------------
G_, Y_, O_, R_, X_ = ('C6EFCE', '006100'), ('FFEB9C', '7F6000'), ('F8CBAD', '843C0C'), ('FFC7CE', '9C0006'), ('E7E6E6', '595959')
DARK = ('C00000', 'FFFFFF')
rules = {
    'Dentro de límite': G_, 'Seguimiento': Y_, 'Alerta': O_, 'Al límite o excedido': R_, 'Límite cero': DARK,
    'Sin información de riesgo': X_, 'Sin límite informado': X_,
    'Aumentar': G_, 'Alargar': G_, 'Reducir': R_, 'Acortar': R_, 'Mantener': Y_, 'Sin definir': X_,
    'Compra conjunta': G_, 'Sesgo comprador': G_, 'Venta conjunta': R_, 'Sesgo vendedor': R_, 'Conflicto': DARK,
    'Información incompleta': X_, 'Todos alargan': G_, 'Sesgo a alargar': G_, 'Todos acortan': R_, 'Sesgo a acortar': R_,
    'Todos mantienen': Y_, 'Conflicto: unos alargan y otros acortan': DARK,
}


def cf(ws, rng_, words):
    for w in words:
        bg, fg = rules[w]
        ws.conditional_formatting.add(rng_, CellIsRule(operator='equal', formula=[f'"{w}"'],
                                                       fill=PatternFill('solid', bgColor=bg, fgColor=bg), font=Font(name=FN, bold=True, color=fg)))


RISK_W = ['Dentro de límite', 'Seguimiento', 'Alerta', 'Al límite o excedido', 'Límite cero', 'Sin información de riesgo', 'Sin límite informado']
DIR_W = ['Aumentar', 'Alargar', 'Reducir', 'Acortar', 'Mantener', 'Sin definir']
CON_W = ['Compra conjunta', 'Sesgo comprador', 'Venta conjunta', 'Sesgo vendedor', 'Conflicto', 'Información incompleta', 'Mantener']
cf(D, f"K{RT + 2}:L{RT + 9}", RISK_W)
cf(D, f"W{RT + 2}:W{RT + 9}", DIR_W)
cf(D, f"Q{CT + 2}:R{CT + 7}", DIR_W)
cf(D, f"F{CT + 8}:L{CT + 8}", ['Todos alargan', 'Sesgo a alargar', 'Todos acortan', 'Sesgo a acortar', 'Todos mantienen',
                               'Conflicto: unos alargan y otros acortan', 'Información incompleta'])
cf(D, f"K{AT + 2}:M{AT + 1 + NCOM}", CON_W)
cf(D, f"N{AT + 2}:Y{AT + 1 + NCOM}", DIR_W)
cf(D, f"B{AT + NCOM + 2}:V{AT + NCOM + 2}", CON_W)
cf(ws_rc, f"{rc['EstadoRiesgo']}{R0}:{rc['EstadoRiesgo']}{R0 + N_PE - 1}", RISK_W)
cf(ws_ac, f"{ac['Consenso']}{R0}:{ac['Consenso']}{R0 + N_ACT - 1}", CON_W)
cf(ws_m, f"{mc['EstadoRiesgo']}{R0}:{mc['EstadoRiesgo']}{R0 + N_POS - 1}", RISK_W)
cf(ws_m, f"{mc['DireccionEstrategia']}{R0}:{mc['DireccionEstrategia']}{R0 + N_POS - 1}", DIR_W)
ws_q.conditional_formatting.add(f"D{QR1}:D{QR2}", CellIsRule(operator='equal', formula=['"Revisar"'],
                                fill=PatternFill('solid', bgColor='FFEB9C', fgColor='FFEB9C'), font=Font(name=FN, bold=True, color='7F6000')))
ws_q.conditional_formatting.add(f"D{QR1}:D{QR2}", CellIsRule(operator='equal', formula=['"OK"'],
                                fill=PatternFill('solid', bgColor='C6EFCE', fgColor='C6EFCE'), font=Font(name=FN, bold=True, color='006100')))

# ---------------- impresión ----------------
D.print_area = f"A1:Z{LAST}"
D.page_setup.orientation = 'landscape'
D.page_setup.paperSize = 8  # A3
D.sheet_properties.pageSetUpPr.fitToPage = True
D.page_setup.fitToWidth = 1
D.page_setup.fitToHeight = 1
D.page_margins.left = D.page_margins.right = D.page_margins.top = D.page_margins.bottom = 0.3
D.print_options.horizontalCentered = True

# =====================================================================
# 09 INSTRUCCIONES
# =====================================================================
I = ws_i
I.sheet_view.showGridLines = False
I.column_dimensions['A'].width = 34
I.column_dimensions['B'].width = 16
I.column_dimensions['C'].width = 13
I.column_dimensions['D'].width = 110
_r = [1]


def H1(t):
    c = I.cell(_r[0], 1, t); c.font = Font(name=FN, size=12, bold=True, color='FFFFFF');
    for j in range(1, 5):
        I.cell(_r[0], j).fill = fill_hdr
    _r[0] += 1


def P(a, b=None, c=None, d=None, bold=False):
    for j, v in enumerate((a, b, c, d), 1):
        if v is not None:
            x = I.cell(_r[0], j, v); x.font = f_bold if (bold or (j == 1 and d is not None)) else f_norm
            x.alignment = Alignment(wrap_text=True, vertical='top')
    _r[0] += 1


def T(t):
    I.merge_cells(start_row=_r[0], start_column=1, end_row=_r[0], end_column=4)
    x = I.cell(_r[0], 1, t); x.font = f_norm; x.alignment = Alignment(wrap_text=True, vertical='top')
    I.row_dimensions[_r[0]].height = 14 * (1 + len(t) // 165)
    _r[0] += 1


I.cell(1, 1, 'Monitor Pensionales: instrucciones').font = Font(name=FN, size=14, bold=True, color=NAVY)
_r[0] = 3
H1('1. Qué hace este libro')
T('Lee tres entradas (posiciones, estrategia y riesgo), las cruza y muestra en 00_DASHBOARD la composición del portafolio elegido, '
  'su exposición por emisor frente al límite, la estrategia y los activos que comparten los portafolios que usted compare.')
T('IMPORTANTE: los datos incluidos son sintéticos (inventados) para demostrar el funcionamiento. Al usar datos reales cambie 04_MAESTROS!B9 a Real.')
_r[0] += 1
H1('2. Actualización diaria')
for t in ['1. Vaya a 01_POSICIONES_INPUT. Seleccione los datos desde la fila 2 hacia abajo (no el encabezado) y bórrelos con la tecla Supr.',
          '2. Pegue la sábana nueva desde la celda A2 como valores, respetando el orden de las columnas del encabezado.',
          '3. Si cambió, haga lo mismo en 02_ESTRATEGIA_INPUT (columnas A a F y H; la columna G es una fórmula informativa y puede quedar vacía).',
          '4. Si llegó archivo de riesgo, haga lo mismo en 03_RIESGO_INPUT.',
          '5. No hay macros ni botón de actualizar: las fórmulas se recalculan solas. Si Excel está en cálculo manual, presione F9.',
          '6. Revise 08_CALIDAD_DATOS: muestra errores y advertencias con la fila de origen.',
          '7. Abra 00_DASHBOARD y elija el portafolio principal, la dimensión de estrategia y hasta 6 portafolios a comparar.',
          'No inserte ni elimine columnas en las entradas. Puede borrar o pegar filas sin problema: el modelo lee por número de fila.']:
    T(t)
_r[0] += 1
H1('3. Código de colores')
P('Verde', d='Aumentar, alargar, compra conjunta, dentro de límite.')
P('Amarillo', d='Mantener o seguimiento (utilización desde 75%). También marca las celdas editables (fondo amarillo suave, letra azul).')
P('Naranja', d='Alerta de riesgo (utilización desde 85%).')
P('Rojo', d='Reducir, acortar, venta conjunta, al límite o excedido (desde 100%). Rojo oscuro: conflicto de estrategia o límite cero con exposición.')
P('Gris', d='Sin información: sin cruce de riesgo, sin límite informado o sin estrategia definida.')
_r[0] += 1
H1('4. Hojas')
for a, d_ in [('00_DASHBOARD', 'Vista de una página. Únicas celdas editables: portafolio principal (B8), dimensión de estrategia (P8) y portafolios a comparar (fila 11).'),
              ('01_POSICIONES_INPUT', 'Entrada diaria: sábana de posiciones (tabla tblPosiciones).'),
              ('02_ESTRATEGIA_INPUT', 'Entrada: estrategia en formato largo (tabla tblEstrategia).'),
              ('03_RIESGO_INPUT', 'Entrada: límites por portafolio y emisor (tabla tblRiesgo).'),
              ('04_MAESTROS', 'Parámetros (umbrales, factor de unidades, modo), orden de nodos, catálogos y equivalencias de NIT.'),
              ('05_MODELO_POSICIONES', 'Posiciones normalizadas con claves, pesos, estrategia y riesgo por fila (tabla tblModeloPosiciones). Solo fórmulas.'),
              ('06_ACTIVOS_COMUNES', 'Un activo por fila con número de portafolios seleccionados que lo tienen y consenso (tabla tblActivosComunes).'),
              ('07_RIESGO_CALCULADO', 'Exposición, límite, utilización, holgura y estado por portafolio y emisor (tabla tblRiesgoCalculado).'),
              ('08_CALIDAD_DATOS', 'Resumen de controles (tabla tblCalidadDatos) y detalle por fila (tabla tblCalidadDetalle).'),
              ('10_CALCULOS', 'Hoja oculta de ayudas del dashboard (rankings y listas para los gráficos). Para verla: clic derecho en una pestaña > Mostrar. No editar.')]:
    P(a, d=d_)
_r[0] += 1
H1('5. Diccionario de datos de las entradas')
P('Campo', 'Tabla', 'Obligatorio', 'Descripción', bold=True)
dic = [
    ('FechaCorte', 'tblPosiciones', 'Sí', 'Fecha de la sábana.'),
    ('CodigoPortafolio', 'tblPosiciones', 'Sí', 'Código estable del portafolio. Se trata como texto.'),
    ('NombrePortafolio', 'tblPosiciones', 'Sí', 'Nombre legible.'),
    ('TipoPortafolio', 'tblPosiciones', 'Sí', 'Solo las filas con el valor Pensional entran al modelo. Las demás se conservan y se reportan como excluidas.'),
    ('ISIN', 'tblPosiciones', 'No', 'Identificador del título. Si falta se construye una clave alternativa.'),
    ('Nemotecnico', 'tblPosiciones', 'Recomendado', 'Nombre corto de mercado. Necesario para la clave alternativa.'),
    ('NITEmisor', 'tblPosiciones', 'Sí', 'NIT como texto. Es la llave del cruce con riesgo.'),
    ('Emisor', 'tblPosiciones', 'Sí', 'Nombre del emisor.'),
    ('Sector', 'tblPosiciones', 'Sí', 'Sector. Vacío se muestra como Sin clasificar.'),
    ('TipoActivo', 'tblPosiciones', 'Sí', 'Categoría de activo (por ejemplo TES tasa fija). Es la dimensión Activo de la estrategia.'),
    ('TipoPapel', 'tblPosiciones', 'Sí', 'TES, Bono ordinario, CDT, FIC, etc.'),
    ('Indicador', 'tblPosiciones', 'Sí', 'Tasa fija, IPC, UVR, IBR, DTF, Liquidez, Otro.'),
    ('Nodo', 'tblPosiciones', 'Sí', 'Rango de plazo. Debe coincidir con la lista de 04_MAESTROS para conservar el orden.'),
    ('FechaVencimiento', 'tblPosiciones', 'No', 'Vencimiento. Se usa en la clave alternativa.'),
    ('ValorPresente', 'tblPosiciones', 'Sí', 'Valoración actual. Unidad del ejemplo: COP MM (ver factor en 04_MAESTROS!B8).'),
    ('ValorNominal', 'tblPosiciones', 'No', 'Nominal. Informativo.'),
    ('Duracion', 'tblPosiciones', 'No', 'Duración de la posición. Se usa para la duración promedio ponderada del portafolio.'),
    ('TasaMercado', 'tblPosiciones', 'No', 'Tasa de referencia. Informativo.'),
    ('Moneda', 'tblPosiciones', 'Sí', 'COP, UVR, USD, etc.'),
    ('Fuente', 'tblPosiciones', 'Sí', 'Archivo o sistema de origen.'),
    ('DimensionEstrategia', 'tblEstrategia', 'Sí', 'Activo, Nodo, Indicador, TipoPapel, Sector, Emisor o Duracion.'),
    ('CategoriaEstrategia', 'tblEstrategia', 'Sí', 'Debe escribirse igual que en la sábana (por ejemplo TES UVR). Para Duracion use Total portafolio.'),
    ('ValorEstrategia', 'tblEstrategia', 'No', 'Positivo aumentar o alargar, negativo reducir o acortar, cero mantener, vacío sin definir.'),
    ('UnidadEstrategia', 'tblEstrategia', 'Sí', 'PuntosPorcentuales (2 = +2 puntos del portafolio), COP_MM o Duracion.'),
    ('DireccionEstrategia', 'tblEstrategia', 'Fórmula', 'Informativa. El modelo calcula la dirección por su cuenta a partir de ValorEstrategia.'),
    ('NITEmisor', 'tblRiesgo', 'Sí', 'Debe coincidir con el NIT de la sábana o tener equivalencia en 04_MAESTROS.'),
    ('LimitePorcentual', 'tblRiesgo', 'Si aplica', 'Fracción del valor total del portafolio (0.10 = 10%).'),
    ('LimiteMonto', 'tblRiesgo', 'Si aplica', 'Monto en la misma unidad de ValorPresente. Tiene prioridad sobre el porcentaje.'),
    ('CupoAprobado, Calificacion', 'tblRiesgo', 'No', 'Informativos.'),
]
for row in dic:
    P(*row)
_r[0] += 1
H1('6. Reglas de cálculo')
for t in ['ClaveActivo: ISIN si existe y tiene la longitud válida; si no, NIT|Nemotécnico|Vencimiento(AAAAMMDD)|Moneda; si tampoco hay vencimiento, NIT|Nemotécnico|TipoPapel|Indicador.',
          'Límite: si hay LimiteMonto se usa; si no, LimitePorcentual por el valor total del portafolio; si no hay ninguno se muestra Sin límite informado. Si no existe fila de riesgo: Sin información de riesgo. Nunca se supone un límite.',
          'Utilización = exposición del emisor en el portafolio / límite. Estados: menor a 75% Dentro de límite; 75% a 85% Seguimiento; 85% a 100% Alerta; 100% o más Al límite o excedido. Umbrales en 04_MAESTROS.',
          'Activo común: la misma ClaveActivo en 2 o más de los portafolios seleccionados. Varias filas del mismo activo en un portafolio cuentan una sola vez.',
          'Consenso: Compra conjunta (todos aumentan), Venta conjunta (todos reducen), Mantener (todos mantienen), Sesgo comprador (aumentar y mantener), Sesgo vendedor (reducir y mantener), Conflicto (al menos uno aumenta y otro reduce), Información incompleta (falta estrategia en alguno).',
          'La dirección de cada activo se toma de la estrategia de su categoría en la dimensión elegida en el dashboard (por defecto Activo, es decir TipoActivo).',
          'Peso objetivo = peso actual + estrategia. Solo se calcula si la unidad es PuntosPorcentuales o COP_MM.']:
    T(t)
_r[0] += 1
H1('7. Supuestos (validar antes de usar con datos reales)')
for t in ['Todos los datos del ejemplo son sintéticos. Portafolios, emisores, NIT, ISIN, montos, límites y estrategias son inventados.',
          'ValorPresente del ejemplo está en COP MM. LimiteMonto y la estrategia en COP_MM usan la misma base.',
          'LimitePorcentual se interpreta como porcentaje del valor total del portafolio. Un valor mayor que 1 (por ejemplo 10) se interpreta como 10%.',
          'Si hay Conflicto y también falta estrategia en algún portafolio, se muestra Conflicto (tiene prioridad sobre Información incompleta).',
          'La composición por emisor agrupa por NIT. Los gráficos muestran las categorías principales y agrupan el resto en Otros. Posiciones negativas quedan dentro de Otros.',
          'En el ejemplo la fecha de riesgo es distinta a la de posiciones a propósito, para mostrar el control de calidad.',
          'Se validan ISIN de 12 caracteres. Si el identificador real tiene otra longitud, cambie 04_MAESTROS!B7 (0 desactiva la validación).',
          'Estrategia: 2 en PuntosPorcentuales significa +2 puntos porcentuales del portafolio.']:
    T(t)
_r[0] += 1
H1('8. Tecnología')
T('Opción B de la especificación: tablas estructuradas y fórmulas clásicas (SUMIFS, COUNTIFS, INDEX, MATCH, SUMPRODUCT). Sin macros, sin Power Query y sin funciones '
  'exclusivas de Microsoft 365, para que abra igual en cualquier versión de Excel. Selectores con listas desplegables (validación de datos).')
_r[0] += 1
H1('9. No implementado o pendiente')
for t in ['Segmentadores (slicers) de fecha, sector e indicador: requieren tablas dinámicas. Se reemplazaron por listas desplegables.',
          'Carga automática desde los archivos reales PlanTrading y Cupos por portafolio: falta el mapeo de sus columnas a estas tres tablas.',
          'Histórico de varias fechas: el libro trabaja con un corte a la vez.',
          'La comparación admite hasta 6 portafolios a la vez. La hoja 06_ACTIVOS_COMUNES también informa en cuántos portafolios Pensionales en total está cada activo.',
          'Derivados, posiciones negativas y liquidez no tienen tratamiento especial.',
          f'Capacidades: {N_POS} posiciones, {N_STR} líneas de estrategia, {N_RSK} líneas de riesgo, {N_ACT} activos distintos, {N_PE} combinaciones portafolio-emisor, '
          f'{N_EMI} emisores, {N_PORT} portafolios. Si se exceden, 08_CALIDAD_DATOS lo avisa como Error crítico.']:
    T(t)

# pestañas y formatos menores
for w, col in ((ws_ma, '7F7F7F'), (ws_m, '7F7F7F'), (ws_ac, TEAL), (ws_rc, TEAL), (ws_q, 'C00000'), (ws_i, '70AD47'), (ws_c, 'BFBFBF')):
    w.sheet_properties.tabColor = col
for w in (ws_ma, ws_q):
    for row in w.iter_rows(min_row=1, max_row=3):
        pass

ws_c.sheet_state = 'hidden'
