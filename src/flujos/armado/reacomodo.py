"""Reemplaza los movimientos de una hoja de banco insertando o quitando filas sin perder estilos."""
from __future__ import annotations

import copy
import logging
import re
from collections.abc import Sequence

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from flujos.armado.estructura import Bloque, Estructura, Fila
from flujos.armado.formulas import MapaFilas

log = logging.getLogger(__name__)
FILAS_MINIMAS_SECCION_VACIA = 3


def _es_formula(v: object) -> bool:
    return isinstance(v, str) and v.startswith('=')


def reemplazar_bloques(ws: Worksheet, e: Estructura, contenido: Sequence[list[Fila]]) -> MapaFilas:
    """Escribe en cada bloque sus filas nuevas y devuelve cómo se movió cada fila anterior.

    `contenido[i]` son los movimientos del bloque e.bloques[i]. El subtotal queda en la última fila del bloque,
    igual que en las hojas actuales, y conserva cualquier otra fórmula que tuviera esa fila.
    """
    bloques = list(e.bloques)
    largos = []
    for b in bloques:  # =SUM(C38) en una sección de una fila: debe crecer con la sección
        rango = rf'SUM({{}}{b.primera}:{{}}{b.ultima})'
        for (f, _), celda in ws._cells.items():
            if f == b.ultima and _es_formula(celda.value):
                celda.value = re.sub(rf'SUM\((\$?[A-Z]+\$?){b.ultima}\)',
                                     lambda m, plantilla=rango: plantilla.format(m.group(1), m.group(1)), celda.value)
    for b, filas in zip(bloques, contenido, strict=True):
        actual = b.ultima - b.primera + 1
        # una sección vacía conserva su espacio; si no, tantas filas como movimientos
        largos.append(max(len(filas), 1) if filas else min(actual, FILAS_MINIMAS_SECCION_VACIA))

    # plan de filas: desplazamiento acumulado al terminar cada bloque
    tramos: list[tuple[Bloque, int, int]] = []  # (bloque, nueva primera, nueva última)
    delta = 0
    for b, largo in zip(bloques, largos, strict=True):
        ini = b.primera + delta
        fin = ini + largo - 1
        tramos.append((b, ini, fin))
        delta += largo - (b.ultima - b.primera + 1)

    def mapa(fila: int, lado: str = 'una') -> int:
        acumulado = 0
        for b, ini, fin in tramos:
            if fila < b.primera:
                return fila + acumulado
            if fila <= b.ultima:
                # un rango abarca el bloque completo; una celda suelta es el subtotal o una fila de movimiento
                if lado == 'ini':
                    return ini
                if lado == 'fin' or fila == b.ultima:
                    return fin
                return ini
            acumulado = fin - b.ultima
        return fila + acumulado

    celdas = dict(ws._cells)
    alturas = {f: copy.copy(d) for f, d in ws.row_dimensions.items()}
    columnas = {c for _, c in celdas}
    en_bloque = {f: i for i, b in enumerate(bloques) for f in range(b.primera, b.ultima + 1)}
    nuevas: dict[tuple[int, int], Cell] = {}

    # 1. lo que no es parte de un bloque se traslada tal cual
    for (f, c), celda in celdas.items():
        if f in en_bloque:
            continue
        celda.row = mapa(f)
        nuevas[(celda.row, c)] = celda

    # 2. cada bloque se reconstruye con los estilos de sus filas actuales
    extras_perdidos = 0
    for i, (b, ini, fin) in enumerate(tramos):
        fila_normal = b.primera
        for n, f in enumerate(range(ini, fin + 1)):
            base = b.ultima if f == fin else fila_normal
            for c in columnas:
                origen = celdas.get((base, c))
                if origen is None:
                    continue
                nueva = Cell(ws, row=f, column=c)
                nueva._style = copy.copy(origen._style)
                nuevas[(f, c)] = nueva
            if f == fin:  # conserva subtotal y demás fórmulas de la última fila
                for c in columnas:
                    origen = celdas.get((b.ultima, c))
                    if origen is not None and _es_formula(origen.value):
                        nuevas[(f, c)].value = origen.value
            filas = contenido[i]
            if n < len(filas):
                fila = filas[n]
                valores = {e.col_fecha: fila.fecha, e.col_concepto: fila.concepto, e.col_debe: fila.debe,
                           e.col_haber: fila.haber, **fila.otros}
                for c, v in valores.items():
                    if v is None:
                        continue
                    if (f, c) not in nuevas:
                        nuevas[(f, c)] = Cell(ws, row=f, column=c)
                    if nuevas[(f, c)].value is None or not _es_formula(nuevas[(f, c)].value):
                        nuevas[(f, c)].value = v
        for f in range(b.primera, b.ultima + 1):
            for c in columnas:
                origen = celdas.get((f, c))
                if origen is not None and origen.value is not None and c > e.col_haber + 1 and not _es_formula(origen.value):
                    extras_perdidos += 1
    if extras_perdidos:
        log.info('%s: %d anotaciones al costado de los movimientos no pasan al mes nuevo.', ws.title, extras_perdidos)

    ws._cells = nuevas
    # 3. alturas de fila y celdas combinadas
    ws.row_dimensions.clear()
    for f, d in alturas.items():
        nueva_fila = mapa(f)
        d.index = nueva_fila
        ws.row_dimensions[nueva_fila] = d
    for rango in list(ws.merged_cells.ranges):
        rango.shift(row_shift=mapa(rango.min_row) - rango.min_row)
    if ws.auto_filter.ref:
        ws.auto_filter.ref = None
    return mapa
