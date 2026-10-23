"""Funciones compartidas: normalización de textos, banco, moneda y nombres de categoría."""
import re, unicodedata

EMPRESAS = ['REVO', 'ECSA', 'ESSA', 'SBSA', 'ONE', 'ICP', 'ELSA', 'ECNA', 'T1']


def norm(s):
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'\s+', ' ', s).strip()


def empresa_de(nombre_archivo):
    n = norm(nombre_archivo)
    for e in EMPRESAS:
        if re.search(r'\b' + e + r'\b', n):
            return e
    return None


def banco(hoja):
    h = norm(hoja)
    for b, ks in [('BCP', ['BCP']), ('BBVA', ['BBVA']), ('SCOT', ['SCOT', 'SBK']), ('IBK', ['IBK']),
                  ('BIF', ['BIF']), ('SANT', ['SANT']), ('PICH', ['PICH']), ('NACION', ['NAC', 'BN']),
                  ('GNB', ['GNB']), ('BCI', ['BCI']), ('FM', ['FM', 'FONDO', 'F.MUTUO'])]:
        if any(re.search(r'(^|[^A-Z])' + re.escape(k), h) or h.startswith(k) for k in ks):
            return b
    return 'OTRO'


def moneda(hoja, cabecera=''):
    """USD o PEN. Primero el nombre de la hoja (es lo más confiable), luego el texto de cabecera."""
    h = norm(hoja)
    if '$' in h or 'DOL' in h or re.search(r'(^|\s)ME$', h):
        return 'USD'
    if (re.search(r'(SOL|PEN|MN\b|NAC|BN\b|CAJA)', h) or re.search(r'(BCP|BBVA|SBK|SCOT|IBK|BIF|SANTANDER)\s?S\b', h)
            or re.search(r'BCP-0', h)):
        return 'PEN'
    c = norm(cabecera)
    if 'M.E' in c or ' ME ' in c or '$' in c or 'DOLAR' in c:
        return 'USD'
    if 'M.N' in c or ' MN ' in c or 'SOLES' in c:
        return 'PEN'
    return 'USD' if banco(hoja) in ('PICH', 'SANT', 'GNB', 'BCI') else 'PEN'


# Variantes de escritura de una misma sección -> un solo nombre (se respeta el estilo de sus hojas).
_FIX = {
    'ABONO TRANSF CTAS': 'ABONO TRANSF. CTAS', 'ABONO TRANSF.': 'ABONO TRANSF. CTAS',
    'CARGO TRANSF CTAS': 'CARGO TRANSF. CTAS', 'CARGO TRASNF CTAS': 'CARGO TRANSF. CTAS',
    'CARGO TRANS. CTAS': 'CARGO TRANSF. CTAS',
    'TRANSF CIAS': 'TRANSF. ENTRE CIAS', 'TRANSF. CIAS': 'TRANSF. ENTRE CIAS', 'TRANSF ENTRE CIAS': 'TRANSF. ENTRE CIAS',
    'TRANSF. ENTRE CIAS.': 'TRANSF. ENTRE CIAS', 'CARGO TRANSF. CIAS': 'TRANSF. ENTRE CIAS',
    'TRANSF CIAS TCP': 'TRANSF. ENTRE CIAS',
    'TRANSF EXTERIOR': 'TRANSF. EXTERIOR', 'TRANSF. EXT.': 'TRANSF. EXTERIOR',
    'TRANSF EXTERIOR (FLETES Y RPTOS)': 'TRANSF. EXTERIOR',
    'PAGO DE DETRACCIONES': 'PAGO DETRACCIONES', 'INGESOS POR ALQUILER': 'INGRESOS POR ALQUILER',
    'DEVOLUCIONES - OTROS EGRESOS': 'DEVOLUCIONES / OTROS EGRESOS',
    'DEVOLUCIONES / OTROS EGRESO': 'DEVOLUCIONES / OTROS EGRESOS',
    'OTROS EGRESOS - DEVOLUCIONES': 'DEVOLUCIONES / OTROS EGRESOS',
    'PAGO GRATIFICACION': 'GRATIFICACIONES', 'PAGO GRATIFICACIONES': 'GRATIFICACIONES',
    'PRESTAMO TERCEROS': 'PRESTAMOS DE TERCEROS', 'FINANCIAMIENTOS': 'FINANCIAMIENTO',
    'OTROS INGRESOS': 'OTROS INGRESOS', 'DUAS': 'PAGO DUAS', 'AFP': 'PAGO AFP',
}


def categoria(seccion):
    s = norm(seccion)
    return _FIX.get(s, s)


# Dentro de estas secciones los movimientos se ordenan por tipo de operación y luego por fecha.
SECCIONES_POR_OPERACION = {'ABONO TRANSF. CTAS', 'CARGO TRANSF. CTAS'}
# (nombre, patrón) en el orden en que aparecen en la hoja; lo que no calza va al final como OTROS.
TIPOS_OPERACION = [
    ('OVERNIGHT', r'\bOV\w*N\w*G\w*T\b|\bOVERNIGHT|\bOVERN'),  # incluye OVERNIGTH, OVENRIGHT
    ('BCR', r'\bBCR\b'),
    ('CAMBIO DE MONEDA', r'DOLARES|CAMBIO (DE )?MONEDA|\bCOMU\$|\bT\.?C\.? ?\d'),
    ('FONDOS MUTUOS', r'FONDOS? MUTUOS?|RESCATE|\bF\.? ?MUTUO|CPDADM'),
]


def tipo_operacion(concepto):
    """Posición del tipo de operación (0 = OVERNIGHT ... len = OTROS) para ordenar las transferencias entre cuentas."""
    c = norm(concepto)
    return next((i for i, (_, p) in enumerate(TIPOS_OPERACION) if re.search(p, c)), len(TIPOS_OPERACION))
