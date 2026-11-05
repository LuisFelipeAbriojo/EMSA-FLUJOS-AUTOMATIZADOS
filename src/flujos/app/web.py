"""Puente entre la página web y el paquete. Corre dentro del navegador con Pyodide.

Cada función recibe rutas del sistema de archivos virtual del navegador y deja lo que hay que descargar
en /salida. Nada se envía a ningún servidor.
"""
from __future__ import annotations

import logging
import os
import shutil
import warnings
import zipfile
from collections.abc import Callable, Iterable
from contextlib import contextmanager
from typing import Any

import pandas as pd

from flujos.armado.proceso import armar_flujo
from flujos.clasificacion.clasificar import clasificar as _clasificar
from flujos.clasificacion.entrenar import entrenar
from flujos.clasificacion.modelo import Clasificador, cargar_modelo, guardar_modelo
from flujos.historico.actualizar import actualizar as _actualizar
from flujos.lectura.normalizar import empresa_de

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

SALIDA = '/salida'
MODELO: Clasificador | None = None
HIST: pd.DataFrame | None = None
Aviso = Callable[[str], Any]


class _ManejadorAviso(logging.Handler):
    def __init__(self, avisar: Aviso) -> None:
        super().__init__(logging.INFO)
        self.avisar = avisar

    def emit(self, record: logging.LogRecord) -> None:
        self.avisar(record.getMessage())


@contextmanager
def _con_avisos(avisar: Aviso):
    """Mientras dura el paso, lo que el paquete registra con logging llega al detalle de la página."""
    raiz = logging.getLogger('flujos')
    nivel = raiz.level
    manejador = _ManejadorAviso(avisar)
    raiz.setLevel(logging.INFO)
    raiz.addHandler(manejador)
    try:
        yield
    finally:
        raiz.removeHandler(manejador)
        raiz.setLevel(nivel)


def _vaciar_salida(ruta: str = SALIDA) -> None:
    shutil.rmtree(ruta, ignore_errors=True)
    os.makedirs(ruta)


def _info_modelo() -> dict[str, Any]:
    m = MODELO
    return dict(movimientos=int(m.n_entrenamiento),
                empresas=sorted({e for e, _ in m.permitidas}),
                categorias=int(len(m.lr.classes_)),
                meses=sorted(HIST.periodo.astype(str).unique().tolist()) if HIST is not None else [])


def cargar_historico(ruta: str, avisar: Aviso) -> dict[str, Any]:
    """Entrena con historico.csv y deja modelo_web.pkl para descargar."""
    global MODELO, HIST
    _vaciar_salida()
    HIST = pd.read_csv(ruta)
    avisar(f'Histórico leído: {len(HIST):,} movimientos. Entrenando (puede tardar unos minutos)...')
    MODELO, _ = entrenar(HIST)
    guardar_modelo(MODELO, f'{SALIDA}/modelo_web.pkl')
    avisar('Modelo entrenado.')
    return _info_modelo()


def cargar_modelo_web(ruta: str, avisar: Aviso) -> dict[str, Any]:
    global MODELO
    _vaciar_salida()
    MODELO = cargar_modelo(ruta)
    avisar('Modelo cargado.')
    return _info_modelo()


def clasificar(rutas: Iterable[str], anio: int, mes: int, umbral: float, avisar: Aviso) -> list[dict[str, Any]]:
    _vaciar_salida()
    rutas = [r for r in rutas if not os.path.basename(r).startswith('~$')]
    with _con_avisos(avisar):
        hechos = _clasificar(rutas, int(anio), int(mes), MODELO, SALIDA, float(umbral))
    for h in hechos:
        h['archivo'] = os.path.basename(h.pop('ruta'))
    if len(hechos) > 1:
        with zipfile.ZipFile(f'{SALIDA}/FLUJOS CLASIFICADOS {anio}-{int(mes):02d}.zip', 'w', zipfile.ZIP_DEFLATED) as z:
            for h in hechos:
                z.write(f"{SALIDA}/{h['archivo']}", h['archivo'])
    return hechos


def actualizar(rutas: Iterable[str], ruta_historico: str | None, umbral: float, avisar: Aviso) -> dict[str, Any]:
    """Rehace los Excel revisados y devuelve el histórico con esos meses agregados."""
    global HIST
    _vaciar_salida()
    if ruta_historico:
        HIST = pd.read_csv(ruta_historico)
    if HIST is None:
        raise ValueError('Falta historico.csv: cárgalo en el paso 1 o junto con los Excel revisados.')
    copias = []
    for r in rutas:
        destino = f'{SALIDA}/{os.path.basename(r)}'
        shutil.copy(r, destino)
        copias.append(destino)
    with _con_avisos(avisar):
        HIST = _actualizar(copias, HIST, MODELO, float(umbral))
    HIST.to_csv(f'{SALIDA}/historico.csv', index=False)
    avisar(f'Histórico actualizado: {len(HIST):,} movimientos.')
    return dict(movimientos=len(HIST))


def armar_flujos(rutas: Iterable[str], anio: int, mes: int, tc: float, avisar: Aviso) -> list[dict[str, Any]]:
    """Empareja por empresa cada Excel clasificado con su flujo del mes anterior y arma el flujo del mes."""
    _vaciar_salida()
    clasificados, anteriores = {}, {}
    for r in rutas:
        nombre = os.path.basename(r)
        empresa = empresa_de(nombre)
        if empresa is None:
            avisar(f'No reconozco la empresa en "{nombre}", lo salto.')
        elif 'CLASIFICADO' in nombre.upper():
            clasificados[empresa] = r
        else:
            anteriores[empresa] = r
    salida = []
    with _con_avisos(avisar):
        for empresa, clasificado in sorted(clasificados.items()):
            anterior = anteriores.get(empresa)
            if anterior is None:
                salida.append(dict(empresa=empresa, error=f'Falta el RESUMEN FLUJO del mes anterior de {empresa}.'))
                continue
            try:
                a = armar_flujo(clasificado, anterior, empresa, int(anio), int(mes), float(tc), SALIDA)
            except ValueError as e:
                salida.append(dict(empresa=empresa, error=str(e)))
                continue
            salida.append(dict(empresa=empresa, flujo=os.path.basename(a.flujo), notas=os.path.basename(a.notas),
                               errores=a.errores, avisos=a.avisos, movimientos=a.movimientos, sin_seccion=a.sin_seccion))
    for empresa in sorted(set(anteriores) - set(clasificados)):
        salida.append(dict(empresa=empresa, error=f'Falta el Excel clasificado de {empresa}.'))
    archivos = [x for x in os.listdir(SALIDA) if x.endswith('.xlsx')]
    if len(archivos) > 2:
        with zipfile.ZipFile(f'{SALIDA}/FLUJOS {anio}-{int(mes):02d}.zip', 'w', zipfile.ZIP_DEFLATED) as z:
            for x in archivos:
                z.write(f'{SALIDA}/{x}', x)
    return salida


def reentrenar(avisar: Aviso) -> dict[str, Any]:
    global MODELO
    _vaciar_salida()
    avisar('Reentrenando con el histórico actualizado...')
    MODELO, _ = entrenar(HIST)
    guardar_modelo(MODELO, f'{SALIDA}/modelo_web.pkl')
    avisar('Modelo reentrenado.')
    return _info_modelo()
