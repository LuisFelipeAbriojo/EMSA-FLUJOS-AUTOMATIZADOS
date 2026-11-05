"""Entiende la anatomía de una hoja de banco del flujo: cabecera, secciones y subtotales."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from flujos.lectura.normalizar import categoria, norm

INGRESO, EGRESO = 'INGRESO', 'EGRESO'
_FECHA_TEXTO = re.compile(r'\d{1,2}\s*[-/. ]\s*\w+\s*[-/. ]\s*\d{2,4}$')


@dataclass
class Fila:
    """Un movimiento en una hoja de banco. Si trae DEBE y HABER a la vez, ocupa una sola fila."""

    fecha: Any
    concepto: Any
    debe: float | None = None
    haber: float | None = None
    otros: dict[int, Any] = field(default_factory=dict)  # columnas adicionales de la fila


@dataclass
class Bloque:
    """Una sección: fila de título, filas de datos y la fila que lleva el subtotal."""

    titulo: str
    fila_titulo: int
    primera: int  # primera fila de datos
    ultima: int  # última fila de datos; aquí suele ir el subtotal
    fila_subtotal: int | None
    tipo: str | None  # INGRESO o EGRESO según qué columna suma el subtotal

    @property
    def categoria(self) -> str:
        return categoria(self.titulo)


@dataclass
class Estructura:
    hoja: str
    fila_cabecera: int
    col_fecha: int
    col_concepto: int
    col_debe: int
    col_haber: int
    bloques: list[Bloque]
    fila_totales: int | None


def valor(ws: Worksheet, fila: int, col: int) -> Any:
    """Lee una celda sin crearla: ws.cell() agrega celdas vacías y vuelve lentas las hojas muy anchas."""
    c = ws._cells.get((fila, col))
    return None if c is None else c.value


def _es_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith('=')


def _es_fecha(v: Any) -> bool:
    return isinstance(v, (dt.datetime, dt.date)) or (isinstance(v, str) and bool(_FECHA_TEXTO.match(v.strip())))


def _es_monto_con_formula(v: Any) -> bool:
    """=750000/SET!E3: un monto escrito como fórmula, no un subtotal ni un vínculo a otra celda."""
    return isinstance(v, str) and bool(re.match(r'=\+?\d', v))


def _es_monto(v: Any) -> bool:
    return (isinstance(v, (int, float)) and not isinstance(v, bool)) or _es_formula(v)


def analizar(ws: Worksheet) -> Estructura | None:
    """Devuelve la estructura de la hoja, o None si no tiene cabecera FECHA/CONCEPTO/DEBE/HABER."""
    cab = None
    for f in range(1, 16):
        nombres: dict[str, int] = {}
        for c in range(1, 12):  # si hay dos cuadros lado a lado, manda el primero
            if valor(ws, f, c) is not None:
                nombres.setdefault(norm(valor(ws, f, c)), c)
        if all(k in nombres for k in ('FECHA', 'CONCEPTO', 'DEBE', 'HABER')):
            cab = (f, nombres)
            break
    if cab is None:
        return None
    f0, nombres = cab
    cf, cc, cd, ch = nombres['FECHA'], nombres['CONCEPTO'], nombres['DEBE'], nombres['HABER']
    ultima_hoja = max((f for f, _ in ws._cells), default=f0)
    bloques: list[Bloque] = []
    actual: dict[str, Any] | None = None
    fila_totales = None

    def cerrar(ultima: int | None = None) -> None:
        nonlocal actual
        if actual is not None and actual['datos']:
            datos = actual['datos']
            bloques.append(Bloque(actual['titulo'], actual['fila'], datos[0], ultima or datos[-1], actual['sub'],
                                  actual['tipo']))
        actual = None

    for f in range(f0 + 1, ultima_hoja + 1):
        fecha, concepto = valor(ws, f, cf), valor(ws, f, cc)
        debe, haber = valor(ws, f, cd), valor(ws, f, ch)
        # fila de totales: una suma de DEBE o HABER que abarca desde arriba de las secciones
        if _es_formula(debe) or _es_formula(haber):
            suma = re.match(r'=\+?SUM\(\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)\)', str(debe if _es_formula(debe) else haber))
            primero = bloques[0].ultima if bloques else f
            if suma and int(suma.group(1)) <= primero and f - int(suma.group(2)) <= 6 and bloques:
                fila_totales = f
                break
        titulo = next((valor(ws, f, c) for c in range(1, cf + 1)
                       if isinstance(valor(ws, f, c), str) and valor(ws, f, c).strip()), None)
        es_titulo = (titulo is not None and not _es_formula(titulo) and not _es_fecha(fecha) and concepto is None
                     and not _es_monto(debe) and not _es_monto(haber))
        if es_titulo:
            cerrar()
            actual = dict(titulo=titulo.strip(), fila=f, datos=[], sub=None, tipo=None)
            continue
        if actual is None:
            continue
        subtotal = next((c for c in range(ch + 1, ch + 4) if _es_formula(valor(ws, f, c))), None)
        hay_dato = _es_fecha(fecha) or _es_monto(debe) or _es_monto(haber) or concepto is not None
        if actual['sub'] is None and (subtotal is not None or hay_dato or valor(ws, f, ch + 1) is not None):
            actual['datos'].append(f)
        elif actual['sub'] is None and not hay_dato:
            actual['datos'].append(f)  # fila vacía dentro de la sección
        if subtotal is not None and actual['sub'] is None:
            actual['sub'] = f
            formula = str(valor(ws, f, subtotal))
            m = re.search(r'SUM\(\$?([A-Z]+)\$?\d+', formula)
            if m:
                columna = range_boundaries(f'{m.group(1)}1')[0]
                actual['tipo'] = INGRESO if columna == cd else EGRESO if columna == ch else None
            siguiente = dict(titulo=actual['titulo'], fila=f, datos=[], sub=None, tipo=None)
            cerrar()
            actual = siguiente  # por si la sección sigue con más datos debajo de su subtotal
            continue
    if actual is not None:
        # recorta las filas vacías del final de la última sección sin subtotal
        while actual['datos'] and all(valor(ws, actual['datos'][-1], c) is None for c in (cf, cc, cd, ch)):
            actual['datos'].pop()
        cerrar()
    bloques = [b for b in bloques if _tiene_contenido(ws, b, cf, cc, cd, ch)]
    return Estructura(ws.title, f0, cf, cc, cd, ch, bloques, fila_totales)


def _tiene_contenido(ws: Worksheet, b: Bloque, cf: int, cc: int, cd: int, ch: int) -> bool:
    if b.fila_subtotal is not None:
        return True
    return any(valor(ws, f, c) is not None for f in range(b.primera, b.ultima + 1) for c in (cf, cc, cd, ch))


def compensa_intereses(ws: Worksheet, e: Estructura, b: Bloque) -> bool:
    """ABONO compensa los intereses con una celda en HABER (D45 = +E417): el abono queda bruto en la sección."""
    return _es_formula(valor(ws, b.ultima, e.col_haber)) and not _es_monto_con_formula(valor(ws, b.ultima, e.col_haber))


def interes_va_en_abono(ws: Worksheet, e: Estructura, intereses: Bloque) -> bool:
    """La sección de intereses es solo un total (C375 = +F47): el detalle está en HABER de cada abono."""
    v = valor(ws, intereses.ultima, e.col_debe)
    return _es_formula(v) and not _es_monto_con_formula(v)


def filas_de_bloque(ws: Worksheet, e: Estructura, b: Bloque, valores: Worksheet | None = None) -> list[Fila]:
    """Los movimientos de una sección. Un vínculo en DEBE/HABER no es un movimiento; `=750000/TC` sí lo es.

    Con `valores` (el libro abierto con valores guardados) esos montos se leen ya calculados.
    """
    filas = []
    for f in range(b.primera, b.ultima + 1):
        fecha, concepto = valor(ws, f, e.col_fecha), valor(ws, f, e.col_concepto)
        d, h = valor(ws, f, e.col_debe), valor(ws, f, e.col_haber)
        if valores is not None:
            d = valores.cell(f, e.col_debe).value if _es_monto_con_formula(d) else d
            h = valores.cell(f, e.col_haber).value if _es_monto_con_formula(h) else h
        d = d if _es_monto(d) and (not _es_formula(d) or _es_monto_con_formula(d)) else None
        h = h if _es_monto(h) and (not _es_formula(h) or _es_monto_con_formula(h)) else None
        if fecha is None and concepto is None and d is None and h is None:
            continue
        otros = {c: valor(ws, f, c) for c in range(1, e.col_fecha) if valor(ws, f, c) is not None}
        filas.append(Fila(fecha, concepto, d, h, otros))
    return filas
