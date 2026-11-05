"""Clasifica el mes cerrado a partir de los libros bancos de cada empresa."""
from __future__ import annotations

import datetime as dt
import logging
import os
from collections.abc import Iterable
from typing import Any

import pandas as pd

from flujos.lectura.libros import leer_con_overnights
from flujos.lectura.normalizar import empresa_de
from flujos.reportes.excel_clasificado import MESES, escribir

log = logging.getLogger(__name__)


def mes_anterior() -> tuple[int, int]:
    h = dt.date.today().replace(day=1) - dt.timedelta(days=1)
    return h.year, h.month


def clasificar(archivos: Iterable[str], anio: int, mes: int, modelo: Any, salida: str = '.',
               umbral: float = 0.9) -> list[dict[str, Any]]:
    """Escribe un Excel clasificado por empresa en `salida` y devuelve un resumen por archivo generado."""
    os.makedirs(salida, exist_ok=True)
    log.info('Mes a procesar: %s %s  (umbral de revisión %.0f%%)\n', MESES[mes], anio, umbral * 100)
    hechos = []
    for f in archivos:
        emp = empresa_de(os.path.basename(f))
        if not emp:
            log.warning('  ?  No reconozco la empresa en "%s", lo salto.', os.path.basename(f))
            continue
        mov, cuentas, overnights = leer_con_overnights(f, anio, mes, emp)
        if mov.empty:
            log.warning('  %-5s sin movimientos en %s %s', emp, MESES[mes], anio)
            continue
        pred = modelo.predict(mov)
        mov = pd.concat([mov.reset_index(drop=True), pred], axis=1)
        ruta = os.path.join(salida, f'{emp} FLUJO {MESES[mes]} {anio} CLASIFICADO.xlsx')
        mov = escribir(ruta, emp, anio, mes, mov, cuentas, modelo, umbral, overnights)
        descuadre = cuentas[
            (cuentas.saldo_inicial + cuentas.ingresos - cuentas.egresos - cuentas.saldo_final).abs() > 0.05]
        hechos.append(dict(empresa=emp, ruta=ruta, movimientos=len(mov), revisar=int(mov.revisar.sum()),
                           cuentas=len(cuentas[cuentas.movimientos > 0]), descuadres=len(descuadre)))
        log.info('  %-5s %5d movimientos en %d cuentas | revisar %4d | cuentas con descuadre de saldo: %d -> %s',
                 emp, len(mov), hechos[-1]['cuentas'], hechos[-1]['revisar'], len(descuadre), ruta)
    return hechos
