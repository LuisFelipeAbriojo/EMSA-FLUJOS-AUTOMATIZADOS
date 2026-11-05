"""Línea de comandos para desarrollo y pruebas. Finanzas usa la app web."""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import logging
import os
import tempfile
import warnings

import pandas as pd

from flujos.armado.proceso import armar_flujo
from flujos.clasificacion.clasificar import clasificar, mes_anterior
from flujos.clasificacion.entrenar import entrenar
from flujos.clasificacion.modelo import cargar_modelo, guardar_modelo
from flujos.config import cargar
from flujos.historico.actualizar import actualizar
from flujos.historico.desde_flujos import abrir_zip, incorporar
from flujos.lectura.flujos import extraer
from flujos.reportes.reporte_historico import escribir_reporte

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


def _cmd_armar(a: argparse.Namespace) -> None:
    anio, mes = map(int, a.mes.split('-'))
    r = armar_flujo(a.clasificado, a.anterior, a.empresa, anio, mes, a.tc, a.salida)
    log.info('%s\n%s', r.flujo, r.notas)


def _cmd_historico(a: argparse.Namespace) -> None:
    h = pd.read_csv(a.historico, low_memory=False)
    with tempfile.TemporaryDirectory() as temporal:
        rutas: list[str] = []
        for x in a.flujos:
            if x.lower().endswith('.zip'):
                rutas += abrir_zip(x, os.path.join(temporal, str(len(rutas))))[0]
            elif os.path.isdir(x):
                rutas += sorted(glob.glob(os.path.join(x, '*.xlsx')))
            else:
                rutas.append(x)
        res = incorporar(rutas, h)
    fecha = dt.date.today().isoformat()
    os.makedirs(a.salida, exist_ok=True)
    escribir_reporte(os.path.join(a.salida, f'REPORTE HISTORICO {fecha}.xlsx'), res, fecha)
    for archivo, motivo in res.rechazados:
        log.warning('Rechazado %s: %s', archivo, motivo)
    if not res.incorporados:
        raise ValueError('Ningún flujo entró al histórico. Mira el reporte.')
    # el histórico anterior no se pisa: cada versión queda con su fecha
    res.historico.to_csv(os.path.join(a.salida, f'historico_{fecha}.csv'), index=False)
    log.info('Histórico: %d -> %d movimientos.', res.antes, len(res.historico))
    if a.modelo:
        modelo, _ = entrenar(res.historico)
        guardar_modelo(modelo, a.modelo)
        log.info('Modelo reentrenado -> %s', a.modelo)


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

    p = sub.add_parser('armar', help='arma el flujo del mes de una empresa')
    p.add_argument('clasificado')
    p.add_argument('anterior', help='RESUMEN FLUJO del mes anterior')
    p.add_argument('--empresa', required=True)
    p.add_argument('--mes', required=True, help='AAAA-MM')
    p.add_argument('--tc', type=float, required=True)
    p.add_argument('--salida', default='.')
    p.set_defaults(fn=_cmd_armar)

    p = sub.add_parser('historico', help='actualiza el histórico con flujos corregidos (.zip, carpetas o .xlsx)')
    p.add_argument('flujos', nargs='+')
    p.add_argument('--historico', required=True, help='historico.csv vigente; no se modifica')
    p.add_argument('--salida', default='.')
    p.add_argument('--modelo', help='si se indica, reentrena y guarda el modelo ahí')
    p.set_defaults(fn=_cmd_historico)

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
