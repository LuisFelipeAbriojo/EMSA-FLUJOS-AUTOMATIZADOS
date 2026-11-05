"""Extrae movimientos ya categorizados de los RESUMEN FLUJO / FLUJO CAJA cerrados.

En cada hoja de banco los movimientos están agrupados bajo filas de título (DEPOSITOS, ITF, PAGO PROVEEDORES...):
ese título es la categoría.
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import re
from typing import Any

import openpyxl

from flujos.lectura.normalizar import banco, categoria, empresa_de, moneda, norm

log = logging.getLogger(__name__)

MESES = {'ENERO': 1, 'ENE': 1, 'FEBRERO': 2, 'FEB': 2, 'MARZO': 3, 'MAR': 3, 'ABRIL': 4, 'ABR': 4, 'MAYO': 5, 'MAY': 5,
         'JUNIO': 6, 'JUN': 6, 'JULIO': 7, 'JUL': 7, 'AGOSTO': 8, 'AGO': 8, 'SETIEMBRE': 9, 'SEPTIEMBRE': 9, 'SET': 9,
         'SEP': 9, 'OCTUBRE': 10, 'OCT': 10, 'NOVIEMBRE': 11, 'NOV': 11, 'DICIEMBRE': 12, 'DIC': 12}


def _num(v: Any) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def extraer(ruta: str) -> list[dict[str, Any]]:
    """Devuelve los movimientos categorizados de un flujo cerrado."""
    nombre = os.path.basename(ruta)
    if 'PROYECTADO' in norm(nombre):
        return []
    emp = empresa_de(nombre)
    m = re.search(r'\b(' + '|'.join(sorted(MESES, key=len, reverse=True)) + r')\b', norm(nombre).replace('-', ' '))
    a = re.search(r'20\d\d', nombre)
    periodo = f"{a.group(0) if a else ''}-{MESES[m.group(1)]:02d}" if m else ''
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    out = []
    omitidas = 0
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(max_col=12, values_only=True))
        cab = None
        for i, r in enumerate(filas[:15]):
            n = [norm(x) if x is not None else '' for x in r]
            if 'FECHA' in n and 'CONCEPTO' in n and 'DEBE' in n and 'HABER' in n:
                cab = (i, n)
                break
        if not cab:
            continue
        i0, n = cab
        cf, cc, cd, ch = n.index('FECHA'), n.index('CONCEPTO'), n.index('DEBE'), n.index('HABER')
        texto_cab = ' '.join(norm(x) for r in filas[:i0] for x in r if x is not None)
        mon, ban = moneda(ws.title, texto_cab), banco(ws.title)
        seccion, orden = None, 0
        for r in filas[i0 + 1:]:
            r = list(r) + [None] * 12
            f, d, h = r[cf], _num(r[cd]), _num(r[ch])
            if isinstance(f, str) and f.strip() and d is None and h is None:
                seccion, orden = categoria(f), orden + 1
                continue
            if not isinstance(f, (dt.datetime, dt.date)) and r[cc] and seccion and (d or h):
                omitidas += 1  # hoy se omiten: fecha escrita como texto o vacía
                continue
            if isinstance(f, (dt.datetime, dt.date)) and r[cc] and seccion and (d or h):
                itf = ('EGRESO' if h < d else 'INGRESO') if d and h and min(d, h) / max(d, h) <= 0.0001 else ''
                for tipo, monto in (('INGRESO', d), ('EGRESO', h)):
                    if monto:
                        out.append(dict(itf_misma_fila=int(tipo == itf), empresa=emp, periodo=periodo, hoja=ws.title,
                                        banco=ban, moneda=mon,
                                        fecha=f.date().isoformat() if isinstance(f, dt.datetime) else f.isoformat(),
                                        concepto=str(r[cc]).strip(), tipo=tipo, monto=round(monto, 2),
                                        categoria=seccion, orden=orden, archivo=nombre))
    if omitidas:
        log.warning('%s: %d filas omitidas por tener la fecha como texto o vacía.', nombre, omitidas)
    return out
