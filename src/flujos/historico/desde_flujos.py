"""Actualiza el histórico con los flujos corregidos y entregados del mes.

El histórico solo aprende de flujos que una persona ya revisó: los Excel clasificados, las notas y los flujos
que salieron del armado sin que nadie los volviera a guardar se rechazan.
"""
from __future__ import annotations

import logging
import os
import zipfile
from dataclasses import dataclass, field
from typing import Any

import openpyxl
import pandas as pd

from flujos.lectura.flujos import FlujoLeido, leer_flujo
from flujos.lectura.normalizar import norm

log = logging.getLogger(__name__)

COLUMNAS = ('empresa', 'periodo', 'hoja', 'moneda', 'fecha', 'concepto', 'tipo', 'monto', 'categoria')
VARIACION_GRANDE = 0.5
MINIMO_EN_PERIODO = 0.5  # con menos, el archivo no es del mes que dice su nombre


@dataclass
class Incorporado:
    archivo: str
    empresa: str
    periodo: str
    accion: str
    antes: int
    ahora: int
    ingresos_pen: float
    egresos_pen: float
    ingresos_usd: float
    egresos_usd: float
    recuperadas: int


@dataclass
class Resultado:
    historico: pd.DataFrame
    antes: int
    incorporados: list[Incorporado] = field(default_factory=list)
    rechazados: list[tuple[str, str]] = field(default_factory=list)  # archivo, motivo
    avisos: list[tuple[str, str]] = field(default_factory=list)  # empresa y periodo, detalle


def validar_historico(h: pd.DataFrame) -> None:
    faltan = [c for c in COLUMNAS if c not in h.columns]
    if faltan:
        raise ValueError(f'El archivo no parece un historico.csv: le faltan las columnas {", ".join(faltan)}.')


def abrir_zip(ruta: str, carpeta: str) -> tuple[list[str], list[str]]:
    """Saca los .xlsx del .zip a `carpeta`, sin respetar sus rutas internas. Devuelve (rutas, nombres ignorados)."""
    try:
        z = zipfile.ZipFile(ruta)
    except zipfile.BadZipFile:
        raise ValueError(f'"{os.path.basename(ruta)}" no es un .zip válido.') from None
    os.makedirs(carpeta, exist_ok=True)
    rutas, ignorados, vistos = [], [], {}
    with z:
        for item in z.infolist():
            if item.is_dir():
                continue
            nombre = os.path.basename(item.filename.replace('\\', '/'))
            if not nombre or nombre.startswith(('~$', '.')) or '__MACOSX' in item.filename:
                continue
            if not nombre.lower().endswith('.xlsx'):
                ignorados.append(nombre)
                continue
            vistos[nombre] = vistos.get(nombre, 0) + 1
            # dos carpetas pueden traer el mismo nombre: no se pisan, y luego se detectan como duplicados
            destino = os.path.join(carpeta, str(vistos[nombre]))
            os.makedirs(destino, exist_ok=True)
            ruta_xlsx = os.path.join(destino, nombre)
            with open(ruta_xlsx, 'wb') as f:
                f.write(z.read(item))
            rutas.append(ruta_xlsx)
    return rutas, ignorados


def _motivo_de_rechazo(nombre: str, flujo: FlujoLeido) -> str | None:
    if flujo.sin_revisar:
        return 'Salió del armado y nadie lo volvió a guardar: revísalo en Excel, guárdalo y vuelve a subirlo.'
    if flujo.empresa is None:
        return 'No reconozco la empresa en el nombre del archivo.'
    if len(flujo.periodo) != 7 or flujo.periodo.startswith('-'):
        return 'No reconozco el mes y el año en el nombre del archivo.'
    if not flujo.movimientos:
        return 'No encontré movimientos con sección. ¿Es un RESUMEN FLUJO cerrado?'
    en_periodo = sum(m['fecha'].startswith(flujo.periodo) for m in flujo.movimientos) / len(flujo.movimientos)
    if en_periodo < MINIMO_EN_PERIODO:
        return f'El nombre dice {flujo.periodo}, pero la mayoría de sus fechas son de otro mes.'
    return None


def _leer(ruta: str) -> tuple[FlujoLeido | None, str | None]:
    nombre = os.path.basename(ruta)
    texto = norm(nombre)
    if 'CLASIFICADO' in texto:
        return None, 'Es un Excel clasificado: el histórico solo aprende de los flujos ya entregados.'
    if 'NOTAS' in texto:
        return None, 'Es un archivo de notas, no un flujo.'
    if 'PROYECTADO' in texto:
        return None, 'Es un proyectado, no un flujo cerrado.'
    try:
        flujo = leer_flujo(ruta)
    except (zipfile.BadZipFile, KeyError, OSError, ValueError, openpyxl.utils.exceptions.InvalidFileException):
        return None, 'No pude abrirlo como Excel (.xlsx). ¿Está dañado?'
    motivo = _motivo_de_rechazo(nombre, flujo)
    return (None, motivo) if motivo else (flujo, None)


def _suma(mov: pd.DataFrame, moneda: str, tipo: str) -> float:
    return round(float(mov[(mov.moneda == moneda) & (mov.tipo == tipo)].monto.sum()), 2)


def _mes_anterior(periodo: str) -> str:
    anio, mes = map(int, periodo.split('-'))
    return f'{anio - 1}-12' if mes == 1 else f'{anio}-{mes - 1:02d}'


def _avisos_del_flujo(flujo: FlujoLeido, mov: pd.DataFrame, h: pd.DataFrame) -> list[tuple[str, str]]:
    clave = f'{flujo.empresa} {flujo.periodo}'
    avisos = [(clave, f'Movimiento sin fecha válida, no entró al histórico. {x}') for x in flujo.omitidas]
    avisos += [(clave, f'Lo leído no suma como la fila de totales (suelen ser ajustes a mano sin concepto). {x}')
               for x in flujo.descuadres]
    if flujo.recuperadas:
        avisos.append((clave, f'{flujo.recuperadas} filas tenían la fecha escrita como texto; se leyeron igual.'))
    fuera = mov[~mov.fecha.str.startswith(flujo.periodo)]
    if not fuera.empty:
        ejemplos = ', '.join(sorted(fuera.fecha.unique())[:5])
        avisos.append((clave, f'{len(fuera)} movimientos con fecha fuera del mes, por ejemplo {ejemplos}.'))
    nuevas = sorted(set(mov.categoria) - set(h.categoria))
    if nuevas:
        avisos.append((clave, f'Secciones que el histórico no conocía: {", ".join(nuevas)}.'))
    previo = h[(h.empresa == flujo.empresa) & (h.periodo == _mes_anterior(flujo.periodo))]
    if len(previo) and abs(len(mov) / len(previo) - 1) > VARIACION_GRANDE:
        avisos.append((clave, f'Tiene {len(mov)} movimientos; el mes anterior tuvo {len(previo)}. Conviene confirmarlo.'))
    return avisos


def incorporar(rutas: list[str], h: pd.DataFrame) -> Resultado:
    """Reemplaza en el histórico cada empresa y periodo que vino en los flujos. No toca lo demás."""
    validar_historico(h)
    res = Resultado(h, len(h))
    leidos: dict[tuple[str, str], FlujoLeido] = {}
    duplicados: set[tuple[str, str]] = set()
    for ruta in sorted(rutas, key=os.path.basename):
        nombre = os.path.basename(ruta)
        flujo, motivo = _leer(ruta)
        if flujo is None:
            res.rechazados.append((nombre, motivo or ''))
            continue
        clave = (flujo.empresa or '', flujo.periodo)
        if clave in leidos or clave in duplicados:
            duplicados.add(clave)
            res.rechazados.append((nombre, f'Vino más de un flujo de {clave[0]} {clave[1]}.'))
            continue
        leidos[clave] = flujo
    for clave in duplicados:
        flujo = leidos.pop(clave)
        res.rechazados.append((flujo.archivo, f'Vino más de un flujo de {clave[0]} {clave[1]}: no sé cuál es el bueno.'))

    pares = pd.Series(list(zip(h.empresa, h.periodo.astype(str), strict=True)), index=h.index, dtype=object)
    conteo = pares.value_counts().to_dict()
    nuevos = []
    for clave, flujo in sorted(leidos.items()):
        mov = pd.DataFrame(flujo.movimientos)
        if 'fuente' in h.columns:
            mov['fuente'] = 'flujo cerrado'
        antes = int(conteo.get(clave, 0))
        res.incorporados.append(Incorporado(
            flujo.archivo, clave[0], clave[1], 'Reemplaza' if antes else 'Nuevo', antes, len(mov),
            _suma(mov, 'PEN', 'INGRESO'), _suma(mov, 'PEN', 'EGRESO'), _suma(mov, 'USD', 'INGRESO'), _suma(mov, 'USD', 'EGRESO'),
            flujo.recuperadas))
        res.avisos += _avisos_del_flujo(flujo, mov, h)
        nuevos.append(mov)
        log.info('%s %s: %d movimientos (%s).', clave[0], clave[1], len(mov), res.incorporados[-1].accion.lower())
    if not nuevos:
        return res

    for periodo in sorted({p for _, p in leidos}):
        faltan = sorted(set(h.empresa) - {e for e, p in leidos if p == periodo})
        if faltan:
            res.avisos.append((periodo, f'No vino el flujo de: {", ".join(faltan)}.'))
    res.historico = pd.concat([h[~pares.isin(set(leidos)).values], *nuevos], ignore_index=True)
    return res


def resumen_numeros(res: Resultado) -> dict[str, Any]:
    return dict(movimientos_antes=res.antes, movimientos=len(res.historico), incorporados=len(res.incorporados),
                rechazados=len(res.rechazados), avisos=len(res.avisos))
