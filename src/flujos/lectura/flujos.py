"""Extrae movimientos ya categorizados de los RESUMEN FLUJO / FLUJO CAJA cerrados.

En cada hoja de banco los movimientos están agrupados bajo filas de título (DEPOSITOS, ITF, PAGO PROVEEDORES...):
ese título es la categoría.
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import re
from dataclasses import dataclass, field
from typing import Any

import openpyxl

from flujos.lectura.libros import leer_fecha
from flujos.lectura.normalizar import banco, categoria, empresa_de, moneda, norm

log = logging.getLogger(__name__)

MESES = {'ENERO': 1, 'ENE': 1, 'FEBRERO': 2, 'FEB': 2, 'MARZO': 3, 'MAR': 3, 'ABRIL': 4, 'ABR': 4, 'MAYO': 5, 'MAY': 5,
         'JUNIO': 6, 'JUN': 6, 'JULIO': 7, 'JUL': 7, 'AGOSTO': 8, 'AGO': 8, 'SETIEMBRE': 9, 'SEPTIEMBRE': 9, 'SET': 9,
         'SEP': 9, 'OCTUBRE': 10, 'OCT': 10, 'NOVIEMBRE': 11, 'NOV': 11, 'DICIEMBRE': 12, 'DIC': 12}

SIN_REVISAR = 'Armado automático, sin revisar'  # Excel lo reemplaza por el nombre de quien guarda el archivo
TOLERANCIA = 0.05
SERIE_MINIMA, SERIE_MAXIMA = 40000, 60000  # números de serie de Excel que son fechas de estos años


@dataclass
class FlujoLeido:
    """Lo que se pudo leer de un flujo cerrado y lo que conviene avisar sobre él."""
    archivo: str
    empresa: str | None
    periodo: str
    movimientos: list[dict[str, Any]] = field(default_factory=list)
    recuperadas: int = 0  # filas con la fecha escrita como texto, leídas igual
    omitidas: list[str] = field(default_factory=list)  # filas con monto que no pudieron entrar
    descuadres: list[str] = field(default_factory=list)  # hojas donde lo leído no suma como su fila de totales
    sin_revisar: bool = False  # salió del armado y nadie lo volvió a guardar


def _num(v: Any) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _fecha(v: Any) -> dt.date | None:
    """Fecha de una celda: real, escrita como texto (11-Set-26, 12Set26) o como número de serie de Excel."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and SERIE_MINIMA <= v <= SERIE_MAXIMA:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    if isinstance(v, str) and v.strip().isdigit():
        return _fecha(int(v.strip()))
    return leer_fecha(v)


def leer_flujo(ruta: str) -> FlujoLeido:
    """Lee los movimientos categorizados de un flujo cerrado."""
    nombre = os.path.basename(ruta)
    emp = empresa_de(nombre)
    m = re.search(r'\b(' + '|'.join(sorted(MESES, key=len, reverse=True)) + r')\b', norm(nombre).replace('-', ' '))
    a = re.search(r'20\d\d', nombre)
    periodo = f"{a.group(0) if a else ''}-{MESES[m.group(1)]:02d}" if m else ''
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    res = FlujoLeido(nombre, emp, periodo, sin_revisar=wb.properties.lastModifiedBy == SIN_REVISAR)
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
        suma_debe, suma_haber, totales = 0.0, 0.0, None
        for r in filas[i0 + 1:]:
            r = list(r) + [None] * 12
            f, d, h = r[cf], _num(r[cd]), _num(r[ch])
            fecha = _fecha(f)
            if isinstance(f, str) and f.strip() and d is None and h is None and fecha is None:
                seccion, orden = categoria(f), orden + 1
                continue
            if totales is None and seccion and fecha is None and not r[cc] and d is not None and h is not None:
                totales = (d, h)
            if not (r[cc] and seccion and (d or h)):
                continue
            if fecha is None:
                if isinstance(r[cc], (int, float)):  # una suma suelta en la columna de concepto, no un movimiento
                    continue
                res.omitidas.append(f'{ws.title}: {str(r[cc]).strip()} ({d or h:,.2f}), sin fecha válida')
                continue
            if not isinstance(f, (dt.datetime, dt.date)):
                res.recuperadas += 1
            itf = ('EGRESO' if h < d else 'INGRESO') if d and h and min(d, h) / max(d, h) <= 0.0001 else ''
            suma_debe, suma_haber = suma_debe + (d or 0), suma_haber + (h or 0)
            for tipo, monto in (('INGRESO', d), ('EGRESO', h)):
                if monto:
                    res.movimientos.append(dict(itf_misma_fila=int(tipo == itf), empresa=emp, periodo=periodo, hoja=ws.title,
                                                banco=ban, moneda=mon, fecha=fecha.isoformat(),
                                                concepto=str(r[cc]).strip(), tipo=tipo, monto=round(monto, 2),
                                                categoria=seccion, orden=orden, archivo=nombre))
        if totales and (abs(totales[0] - suma_debe) > TOLERANCIA or abs(totales[1] - suma_haber) > TOLERANCIA):
            res.descuadres.append(f'{ws.title}: leído {suma_debe:,.2f} / {suma_haber:,.2f}, '
                                  f'la hoja totaliza {totales[0]:,.2f} / {totales[1]:,.2f}')
    if res.omitidas:
        log.warning('%s: %d filas omitidas por no tener una fecha válida.', nombre, len(res.omitidas))
    return res


def extraer(ruta: str) -> list[dict[str, Any]]:
    """Devuelve los movimientos categorizados de un flujo cerrado."""
    return [] if 'PROYECTADO' in norm(os.path.basename(ruta)) else leer_flujo(ruta).movimientos
