"""Datos SINTETICOS de demostracion. No representan posiciones reales."""
import datetime as dt

FC = dt.date(2026, 11, 10)
FR = dt.date(2026, 11, 7)   # fecha de riesgo distinta a proposito (prueba del control)
SRC = 'SINTETICO_DEMO'

HDR_P = ['FechaCorte', 'CodigoPortafolio', 'NombrePortafolio', 'TipoPortafolio', 'ISIN', 'Nemotecnico', 'NITEmisor', 'Emisor',
         'Sector', 'TipoActivo', 'TipoPapel', 'Indicador', 'Nodo', 'FechaVencimiento', 'ValorPresente', 'ValorNominal',
         'Duracion', 'TasaMercado', 'Moneda', 'Fuente']
HDR_E = ['FechaCorte', 'CodigoPortafolio', 'DimensionEstrategia', 'CategoriaEstrategia', 'ValorEstrategia', 'UnidadEstrategia',
         'DireccionEstrategia', 'Comentario']
HDR_R = ['FechaCorte', 'CodigoPortafolio', 'NITEmisor', 'EmisorRiesgo', 'TipoLimite', 'LimitePorcentual', 'LimiteMonto',
         'CupoAprobado', 'Calificacion', 'FuenteRiesgo']

PORT = {
    'PEN001': 'Fondo Pensional Ejemplo Norte',
    'PEN002': 'Fondo Pensional Ejemplo Centro',
    'PEN003': 'Fondo Pensional Ejemplo Sur',
    'PEN004': 'Fondo Pensional Ejemplo Oriente',
}
EMI = {  # clave: (NIT, nombre, sector)
    'GOB': ('899999001', 'Gobierno Nacional Ejemplo', 'Gobierno nacional'),
    'ALF': ('900000002', 'Emisor Financiero Alfa', 'Financiero'),
    'GAM': ('900000003', 'Emisor Financiero Gamma', 'Financiero'),
    'DEL': ('900000004', 'Banco Ejemplo Delta', 'Financiero'),
    'BET': ('900000005', 'Empresa Energía Beta', 'Energía'),
    'EPS': ('900000006', 'Infraestructura Épsilon', 'Infraestructura'),
    'ZET': ('900000007', 'Servicios Públicos Zeta', 'Servicios públicos'),
    'ETA': ('900000008', 'Corporativo Eta', 'Corporativo'),
    'THE': ('900000009', 'Titularizadora Theta', 'Financiero'),
    'IOT': ('900000010', 'Gestora FIC Iota', 'Financiero'),
    'KAP': ('900000011', 'Banco Liquidez Kappa', 'Liquidez'),
}
N1, N2, N3, N4, N5, N6, N7 = ('Liquidez / menor a 1 año', '1-3 años', '3-5 años', '5-7 años', '7-10 años', '10-15 años', 'Mayor a 15 años')
TF, IPC, UVR, IBR, DTF, LIQ, OTR = 'Tasa fija', 'IPC', 'UVR', 'IBR', 'DTF', 'Liquidez', 'Otro'
d = dt.date
# id: (ISIN, nemo, emisor, sector(None=usar el del emisor), TipoActivo, TipoPapel, Indicador, Nodo, venc, dur, tasa, moneda)
A = {
    'A01': ('COSYN0000011', 'TESTF28', 'GOB', None, 'TES tasa fija', 'TES', TF, N2, d(2028, 4, 28), 1.35, 0.0985, 'COP'),
    'A02': ('COSYN0000029', 'TESTF31', 'GOB', None, 'TES tasa fija', 'TES', TF, N3, d(2031, 3, 26), 3.60, 0.1090, 'COP'),
    'A03': ('COSYN0000037', 'TESTF34', 'GOB', None, 'TES tasa fija', 'TES', TF, N5, d(2034, 10, 18), 5.55, 0.1165, 'COP'),
    'A04': ('COSYN0000045', 'TESTF42', 'GOB', None, 'TES tasa fija', 'TES', TF, N7, d(2042, 7, 9), 7.90, 0.1230, 'COP'),
    'A05': ('COSYN0000052', 'TESUVR29', 'GOB', None, 'TES UVR', 'TES', UVR, N2, d(2029, 4, 18), 2.30, 0.0510, 'UVR'),
    'A06': ('COSYN0000060', 'TESUVR35', 'GOB', None, 'TES UVR', 'TES', UVR, N5, d(2035, 5, 9), 7.20, 0.0565, 'UVR'),
    'A07': ('COSYN0000078', 'TESUVR49', 'GOB', None, 'TES UVR', 'TES', UVR, N7, d(2049, 6, 16), 14.10, 0.0590, 'UVR'),
    'A08': ('COSYN0000086', 'BALFIPC30', 'ALF', None, 'Deuda privada IPC', 'Bono ordinario', IPC, N3, d(2030, 8, 14), 3.20, 0.1120, 'COP'),
    'A09': ('COSYN0000094', 'BALFTF28', 'ALF', None, 'Deuda privada tasa fija', 'Bono ordinario', TF, N2, d(2028, 9, 5), 1.65, 0.1075, 'COP'),
    'A10': ('COSYN0000102', 'CDTGAMIBR27', 'GAM', None, 'Deuda privada IBR', 'CDT', IBR, N1, d(2027, 5, 20), 0.20, 0.1010, 'COP'),
    'A11': ('COSYN0000110', 'CDTGAMTF27', 'GAM', None, 'Deuda privada tasa fija', 'CDT', TF, N1, d(2027, 8, 12), 0.72, 0.1040, 'COP'),
    'A12': ('COSYN0000128', 'BSUBDELIPC33', 'DEL', None, 'Deuda privada IPC', 'Bono subordinado', IPC, N4, d(2033, 2, 22), 4.90, 0.1210, 'COP'),
    'A13': ('COSYN0000136', 'CDTDELTF28', 'DEL', None, 'Deuda privada tasa fija', 'CDT', TF, N2, d(2028, 3, 15), 1.25, 0.1060, 'COP'),
    'A14': ('COSYN0000144', 'BBETIPC36', 'BET', None, 'Deuda privada IPC', 'Bono ordinario', IPC, N5, d(2036, 6, 30), 6.40, 0.1195, 'COP'),
    'A15': ('COSYN0000151', 'BBETIPC40', 'BET', None, 'Deuda privada IPC', 'Bono ordinario', IPC, N6, d(2040, 11, 21), 8.30, 0.1240, 'COP'),
    'A16': ('COSYN0000169', 'BEPSUVR38', 'EPS', None, 'Deuda privada UVR', 'Bono ordinario', UVR, N6, d(2038, 9, 10), 8.60, 0.0640, 'UVR'),
    'A17': ('COSYN0000177', 'BZETIPC32', 'ZET', None, 'Deuda privada IPC', 'Bono ordinario', IPC, N4, d(2032, 12, 3), 4.60, 0.1150, 'COP'),
    'A18': (None, 'BETATF29', 'ETA', None, 'Deuda privada tasa fija', 'Bono ordinario', TF, N2, d(2029, 6, 11), 2.25, 0.1180, 'COP'),   # ISIN vacio
    'A19': ('COSYN0000193', 'BETAIBR28', 'ETA', '', 'Deuda privada IBR', 'Bono ordinario', IBR, N2, d(2028, 10, 2), 0.25, 0.1105, 'COP'),  # sector vacio
    'A20': ('COSYN0000201', 'TITTHEUVR37', 'THE', None, 'Deuda privada UVR', 'Titularización', UVR, N6, d(2037, 4, 27), 6.90, 0.0660, 'UVR'),
    'A21': ('COSYN0000219', 'FICIOTA', 'IOT', None, 'FIC', 'FIC', OTR, N1, None, 0.10, None, 'COP'),
    'A22': (None, 'CTAKAPPA', 'KAP', None, 'Liquidez', 'Cuenta bancaria / liquidez', LIQ, N1, None, 0.0, 0.0925, 'COP'),            # sin ISIN ni vencimiento
    'A23': ('COSYN0000235', 'CDTALFDTF27', 'ALF', None, 'Deuda privada DTF', 'CDT', DTF, N1, d(2027, 2, 18), 0.22, 0.1030, 'COP'),
    'A24': ('COSYN0000243', 'TESTF50', 'GOB', None, 'TES tasa fija', 'TES', TF, N7, d(2050, 10, 26), 8.60, 0.1245, 'COP'),
    'A25': ('COSYN0000250', 'BDELTF31', 'DEL', None, 'Deuda privada tasa fija', 'Bono ordinario', TF, N3, d(2031, 7, 7), 3.70, 0.1135, 'COP'),
    'A26': ('COSYN0000268', 'BZETTF30', 'ZET', None, 'Deuda privada tasa fija', 'Bono ordinario', TF, N3, d(2030, 5, 19), 3.00, 0.1110, 'COP'),
}
# tenencias: (portafolio, activo, ValorPresente COP MM)
H = [
    ('PEN001', 'A01', 8000), ('PEN001', 'A02', 7000), ('PEN001', 'A02', 5000), ('PEN001', 'A03', 10000), ('PEN001', 'A05', 7000),
    ('PEN001', 'A06', 9000), ('PEN001', 'A08', 6000), ('PEN001', 'A09', 3200), ('PEN001', 'A10', 4000), ('PEN001', 'A11', 3800),
    ('PEN001', 'A14', 5500), ('PEN001', 'A15', 3000), ('PEN001', 'A18', 2500), ('PEN001', 'A12', 4000), ('PEN001', 'A17', 5000),
    ('PEN001', 'A21', 3000), ('PEN001', 'A22', 6000), ('PEN001', 'A16', 4000), ('PEN001', 'A20', 4000),
    ('PEN002', 'A02', 9000), ('PEN002', 'A04', 6000), ('PEN002', 'A06', 8000), ('PEN002', 'A07', 5000), ('PEN002', 'A08', 4000),
    ('PEN002', 'A10', 3000), ('PEN002', 'A11', 2500), ('PEN002', 'A13', 3500), ('PEN002', 'A14', 4000), ('PEN002', 'A21', 2000),
    ('PEN002', 'A22', 4000), ('PEN002', 'A23', 2000), ('PEN002', 'A25', 3000), ('PEN002', 'A19', 4000),
    ('PEN003', 'A01', 6000), ('PEN003', 'A02', 10000), ('PEN003', 'A05', 8000), ('PEN003', 'A07', 6000), ('PEN003', 'A08', 5000),
    ('PEN003', 'A12', 4000), ('PEN003', 'A15', 5000), ('PEN003', 'A16', 5000), ('PEN003', 'A17', 4000), ('PEN003', 'A20', 6000),
    ('PEN003', 'A22', 5000), ('PEN003', 'A24', 7000), ('PEN003', 'A26', 4000), ('PEN003', 'A09', 5000),
    ('PEN004', 'A03', 7000), ('PEN004', 'A04', 5000), ('PEN004', 'A06', 6000), ('PEN004', 'A13', 3000), ('PEN004', 'A16', 3000),
    ('PEN004', 'A21', 2000), ('PEN004', 'A22', 3000), ('PEN004', 'A24', 5000), ('PEN004', 'A25', 3000), ('PEN004', 'A26', 3000),
]


def _row(pcode, pname, ptype, aid, vp):
    isin, nemo, e, sec, ta, tp, ind, nodo, venc, dur, tasa, mon = A[aid]
    nit, emisor, sector = EMI[e]
    if sec is not None:
        sector = sec or None
    return [FC, pcode, pname, ptype, isin, nemo, nit, emisor, sector, ta, tp, ind, nodo, venc, float(vp), round(vp * 0.97, 2),
            dur, tasa, mon, SRC]


POS = [_row(p, PORT[p], 'Pensional', a, v) for p, a, v in H]
POS.append(_row('VOL001', 'Fondo Voluntario Ejemplo (no Pensional)', 'Voluntario', 'A02', 1000))

PP, MM, DU = 'PuntosPorcentuales', 'COP_MM', 'Duracion'
_E = [
    # Activo (dimension por defecto)
    ('PEN001', 'Activo', 'TES tasa fija', 2.0, PP, 'Aumentar TES tasa fija'),
    ('PEN001', 'Activo', 'TES UVR', -1.5, PP, 'Recortar UVR'),
    ('PEN001', 'Activo', 'Deuda privada tasa fija', -1.0, PP, ''),
    ('PEN001', 'Activo', 'Deuda privada IPC', 1.0, PP, ''),
    ('PEN001', 'Activo', 'Deuda privada IBR', 0.0, PP, 'Cero = decisión explícita de mantener'),
    ('PEN001', 'Activo', 'Liquidez', -0.5, PP, ''),
    ('PEN001', 'Activo', 'FIC', None, PP, 'Vacío = sin estrategia definida (no es cero)'),
    ('PEN002', 'Activo', 'TES tasa fija', 1.5, PP, ''),
    ('PEN002', 'Activo', 'TES UVR', -2.0, PP, ''),
    ('PEN002', 'Activo', 'Deuda privada tasa fija', 0.0, PP, ''),
    ('PEN002', 'Activo', 'Deuda privada IPC', -1.0, PP, ''),
    ('PEN002', 'Activo', 'Deuda privada IBR', 0.0, PP, ''),
    ('PEN002', 'Activo', 'Deuda privada DTF', -0.5, PP, ''),
    ('PEN002', 'Activo', 'Liquidez', 1.0, PP, ''),
    ('PEN002', 'Activo', 'FIC', 0.5, PP, ''),
    ('PEN003', 'Activo', 'TES tasa fija', 1.0, PP, ''),
    ('PEN003', 'Activo', 'TES UVR', 1.0, PP, ''),
    ('PEN003', 'Activo', 'Deuda privada tasa fija', -0.5, PP, ''),
    ('PEN003', 'Activo', 'Deuda privada IPC', 0.5, PP, ''),
    ('PEN003', 'Activo', 'Deuda privada UVR', -1.0, PP, ''),
    ('PEN003', 'Activo', 'Liquidez', 0.0, PP, ''),
    ('PEN004', 'Activo', 'TES tasa fija', 0.0, PP, ''),
    ('PEN004', 'Activo', 'TES UVR', -1.0, PP, ''),
    ('PEN004', 'Activo', 'Deuda privada tasa fija', 0.5, PP, ''),
    ('PEN004', 'Activo', 'Deuda privada UVR', -0.5, PP, ''),
    ('PEN004', 'Activo', 'FIC', 0.0, PP, ''),
    ('PEN004', 'Activo', 'Liquidez', 0.0, PP, ''),
    # Duracion del portafolio (alargar / acortar)
    ('PEN001', 'Duracion', 'Total portafolio', 0.30, DU, 'Alargar 0.30 años'),
    ('PEN002', 'Duracion', 'Total portafolio', 0.20, DU, ''),
    ('PEN003', 'Duracion', 'Total portafolio', -0.25, DU, 'Acortar'),
    ('PEN004', 'Duracion', 'Total portafolio', 0.0, DU, 'Mantener'),
    # Otras dimensiones (para probar el selector de dimension)
    ('PEN001', 'Nodo', N5, 1.0, PP, ''),
    ('PEN001', 'Nodo', N2, -1.0, PP, ''),
    ('PEN002', 'Nodo', N5, 0.5, PP, ''),
    ('PEN003', 'Nodo', N2, 1.0, PP, ''),
    ('PEN001', 'Indicador', IPC, 1.0, PP, ''),
    ('PEN001', 'Indicador', UVR, -1.5, PP, ''),
    ('PEN002', 'Indicador', UVR, -2.0, PP, ''),
    ('PEN001', 'Emisor', 'Empresa Energía Beta', -500.0, MM, 'Reducir emisor excedido (ejemplo en COP MM)'),
]
EST = [(FC, c, dim, cat, val, uni, com) for c, dim, cat, val, uni, com in _E]

PCT, MON = 'Porcentaje del portafolio', 'Monto absoluto'
_R = [
    ('PEN001', 'GOB', PCT, 1.00, None, 'Nación'), ('PEN001', 'ALF', PCT, 0.10, None, 'AAA'), ('PEN001', 'GAM', PCT, 0.10, None, 'AAA'),
    ('PEN001', 'BET', PCT, 0.08, None, 'AA+'), ('PEN001', 'DEL', PCT, 0.10, None, 'AAA'), ('PEN001', 'ZET', MON, None, 9000.0, 'AA+'),
    ('PEN001', 'IOT', PCT, 0.05, None, 'AAA'), ('PEN001', 'KAP', PCT, 0.15, None, 'AAA'), ('PEN001', 'EPS', PCT, 0.06, None, 'AA'),
    ('PEN001', 'THE', PCT, 0.05, None, 'AAA'),                       # ETA sin fila: cruce faltante
    ('PEN002', 'GOB', PCT, 1.00, None, 'Nación'), ('PEN002', 'ALF*', PCT, 0.12, None, 'AAA'), ('PEN002', 'GAM', PCT, 0.10, None, 'AAA'),
    ('PEN002', 'DEL', PCT, 0.10, None, 'AAA'), ('PEN002', 'BET', PCT, 0.08, None, 'AA+'), ('PEN002', 'IOT', PCT, 0.05, None, 'AAA'),
    ('PEN002', 'KAP', PCT, 0.15, None, 'AAA'), ('PEN002', 'ETA', PCT, 0.08, None, 'AA'),
    ('PEN003', 'GOB', PCT, 1.00, None, 'Nación'), ('PEN003', 'ALF', PCT, 0.15, None, 'AAA'), ('PEN003', 'DEL', PCT, 0.10, None, 'AAA'),
    ('PEN003', 'BET', PCT, 0.08, None, 'AA+'), ('PEN003', 'EPS', PCT, 0.06, None, 'AA'), ('PEN003', 'ZET', MON, None, 9000.0, 'AA+'),
    ('PEN003', 'THE', 'Sin definir', None, None, 'AAA'),             # fila sin limite informado
    ('PEN003', 'KAP', PCT, 0.15, None, 'AAA'),
    ('PEN004', 'GOB', PCT, 1.00, None, 'Nación'), ('PEN004', 'DEL', PCT, 0.20, None, 'AAA'), ('PEN004', 'EPS', MON, None, 0.0, 'AA'),
    ('PEN004', 'IOT', PCT, 0.05, None, 'AAA'), ('PEN004', 'KAP', PCT, 0.15, None, 'AAA'), ('PEN004', 'ZET', PCT, 0.10, None, 'AA+'),
]
RSK = []
for c, e, tipo, pct, monto, cal in _R:
    alt = e.endswith('*')
    nit, nom, _ = EMI[e.rstrip('*')]
    RSK.append([FR, c, nit + '-1' if alt else nit, nom, tipo, pct, monto, monto, cal, SRC])
EQUIV = [('900000002-1', '900000002')]
