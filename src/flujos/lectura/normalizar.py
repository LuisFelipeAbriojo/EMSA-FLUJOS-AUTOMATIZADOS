"""Normalización de textos y detección de empresa, banco, moneda y categoría."""
from __future__ import annotations

import re
import unicodedata
from typing import Any

from flujos.config import cargar


def norm(s: Any) -> str:
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'\s+', ' ', s).strip()


def empresa_de(nombre_archivo: str) -> str | None:
    n = norm(nombre_archivo)
    for e in cargar()['empresas']['codigos']:
        if re.search(r'\b' + e + r'\b', n):
            return e
    return None


def banco(hoja: str) -> str:
    h = norm(hoja)
    for b in cargar()['bancos']:
        if any(re.search(r'(^|[^A-Z])' + re.escape(k), h) or h.startswith(k) for k in b['claves']):
            return b['codigo']
    return 'OTRO'


def moneda(hoja: str, cabecera: str = '') -> str:
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
    return 'USD' if banco(hoja) in cargar()['monedas']['bancos_en_dolares_por_defecto'] else 'PEN'


_OVERNIGHT = re.compile(r'OV\w*N\w*G\w*T|OVERN')


def es_overnight(concepto: Any) -> bool:
    return bool(_OVERNIGHT.search(norm(concepto)))


def categoria(seccion: str) -> str:
    """Unifica las variantes de escritura de una sección en un solo nombre."""
    s = norm(seccion)
    return cargar()['categorias']['variantes'].get(s, s)
