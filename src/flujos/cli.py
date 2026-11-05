"""Línea de comandos para desarrollo y pruebas. Finanzas usa la app web."""
from __future__ import annotations

import argparse
import csv
import glob
import logging
import os
import warnings

import pandas as pd

from flujos.clasificacion.clasificar import clasificar, mes_anterior
from flujos.clasificacion.entrenar import entrenar
from flujos.clasificacion.modelo import cargar_modelo, guardar_modelo
from flujos.config import cargar
from flujos.historico.actualizar import actualizar
from flujos.lectura.flujos import extraer

log = logging.getLogger('flujos')


def _cmd_entrenar(a: argparse.Namespace) -> None:
    modelo, h = entrenar(pd.read_csv(a.historico))
    guardar_modelo(modelo, a.modelo)
    log.info('Modelo entrenado con %d movimientos, %d empresas, %d meses, %d categorías -> %s',
             len(h), h.empresa.nunique(), h.periodo.nunique(), h.categoria.nunique(), a.modelo)


def _cmd_clasificar(a: argparse.Namespace) -> None:
    anio, mes = map(int, a.mes.split('-')) if a.mes else mes_anterior()
    archivos: list[str] = []
    for x in a.libros:
        archivos += sorted(glob.glob(os.path.join(x, '*.xlsx'))) if os.path.isdir(x) else [x]
    archivos = [f for f in archivos if not os.path.basename(f).startswith('~$')]
    clasificar(archivos, anio, mes, cargar_modelo(a.modelo), a.salida, a.umbral)


def _cmd_actualizar(a: argparse.Namespace) -> None:
    h = pd.read_csv(a.historico) if os.path.exists(a.historico) else pd.DataFrame()
    actualizar(a.archivos, h, cargar_modelo(a.modelo), a.umbral).to_csv(a.historico, index=False)
    log.info('Histórico actualizado: %s. Ahora vuelve a entrenar el modelo.', a.historico)


def _cmd_extraer(a: argparse.Namespace) -> None:
    filas = []
    for c in a.carpetas:
        for f in sorted(glob.glob(os.path.join(c, '**', '*.xlsx'), recursive=True)):
            r = extraer(f)
            log.info('%6d  %s', len(r), os.path.basename(f))
            filas += r
    if not filas:
        raise SystemExit('No encontré movimientos en esas carpetas.')
    with open(a.salida, 'w', newline='', encoding='utf-8') as fo:
        w = csv.DictWriter(fo, fieldnames=list(filas[0]))
        w.writeheader()
        w.writerows(filas)
    log.info('%d movimientos -> %s', len(filas), a.salida)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog='flujos', description=__doc__)
    sub = ap.add_subparsers(dest='comando', required=True)
    umbral = cargar()['clasificacion']['umbral_revision']

    p = sub.add_parser('entrenar', help='entrena el modelo con el histórico')
    p.add_argument('historico')
    p.add_argument('modelo')
    p.set_defaults(fn=_cmd_entrenar)

    p = sub.add_parser('clasificar', help='clasifica el mes desde los libros bancos')
    p.add_argument('libros', nargs='+', help='archivos .xlsx o carpetas que los contengan')
    p.add_argument('--modelo', required=True)
    p.add_argument('--mes', help='AAAA-MM (por defecto, el mes anterior)')
    p.add_argument('--salida', default='.')
    p.add_argument('--umbral', type=float, default=umbral)
    p.set_defaults(fn=_cmd_clasificar)

    p = sub.add_parser('actualizar', help='rehace los Excel revisados y agrega el mes al histórico')
    p.add_argument('archivos', nargs='+')
    p.add_argument('--historico', required=True)
    p.add_argument('--modelo', required=True)
    p.add_argument('--umbral', type=float, default=umbral)
    p.set_defaults(fn=_cmd_actualizar)

    p = sub.add_parser('extraer', help='arma un histórico desde RESUMEN FLUJO cerrados')
    p.add_argument('salida')
    p.add_argument('carpetas', nargs='+')
    p.set_defaults(fn=_cmd_extraer)

    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    try:
        a.fn(a)
    except (ValueError, FileNotFoundError) as e:
        log.error('Error: %s', e)
        return 1
    return 0
