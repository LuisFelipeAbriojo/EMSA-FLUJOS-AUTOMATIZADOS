"""Después de revisar: rehace las hojas por cuenta del Excel clasificado con las categorías corregidas."""
from __future__ import annotations

import logging
import os
import re
from collections.abc import Iterable
from typing import Any

import openpyxl
import pandas as pd

from flujos.reportes.excel_clasificado import MESES, escribir

log = logging.getLogger(__name__)


def rehacer(archivos: Iterable[str], modelo: Any, umbral: float = 0.9) -> None:
    """Rehace cada archivo revisado en su sitio. El histórico no cambia: solo aprende de los flujos entregados."""
    for f in archivos:
        nombre = os.path.basename(f)
        emp = nombre.split()[0]
        m = re.search(r'FLUJO (\w+) (\d{4})', nombre)
        if not m or m.group(1) not in MESES:
            raise ValueError(f'"{nombre}" no parece un Excel clasificado: se espera un nombre como '
                             f'"ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx".')
        mes, anio = MESES.index(m.group(1)), int(m.group(2))
        mv = pd.read_excel(f, sheet_name='MOVIMIENTOS')
        # saldos y cuentas desde la hoja RESUMEN
        wb = openpyxl.load_workbook(f, data_only=False)
        ws = wb['RESUMEN']
        cuentas = []
        for r in ws.iter_rows(min_row=5, values_only=True):
            if not r[0] or r[2] not in ('USD', 'PEN'):
                break
            cuentas.append(dict(hoja=r[0], cuenta=r[1], moneda=r[2], saldo_inicial=r[3], ingresos=r[4], egresos=r[5],
                                saldo_final=r[7], movimientos=r[9]))
        cuentas = pd.DataFrame(cuentas)
        # las fechas sospechosas siguen siendo del libro: se conservan al rehacer el RESUMEN
        alertas = [c[2:] for c, in ws.iter_rows(max_col=1, values_only=True)
                   if isinstance(c, str) and c.startswith('• ')]
        cuentas['alertas'] = ''
        cuentas.loc[0, 'alertas'] = ' | '.join(alertas)
        mov = pd.DataFrame(dict(
            empresa=emp, hoja=mv.CUENTA, moneda=mv.MONEDA, fecha=pd.to_datetime(mv.FECHA).dt.date,
            concepto=mv.CONCEPTO, tipo=mv.TIPO, monto=mv.MONTO,
            categoria=mv.CATEGORIA.astype(str).str.strip().str.upper(),
            confianza=mv.CONFIANZA, alternativa=mv.ALTERNATIVA, origen=mv.ORIGEN, aviso=mv.NOTA.fillna(''),
            operacion=mv['Nº OPERACION'], cheque=mv['CH/Nº'], fila_origen=mv['FILA EN LIBRO'],
            itf_misma_fila=mv.NOTA.fillna('').str.contains('misma fila').astype(int), concepto_padre=''))
        # lo que el equipo revisó ya no está "por revisar"
        mov.loc[mov.confianza < umbral, ['confianza', 'origen']] = [1.0, 'revisado']
        overnights = pd.read_excel(f, sheet_name='OVERNIGHTS') if 'OVERNIGHTS' in wb.sheetnames else None
        escribir(f, emp, anio, mes, mov, cuentas, modelo, umbral, overnights)
        log.info('%s: hojas por cuenta rehechas con %d movimientos.', emp, len(mov))
