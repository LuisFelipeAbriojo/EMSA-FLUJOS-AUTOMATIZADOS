"""Validaciones automáticas del flujo armado. Cada una dice si pasó y, si no, qué revisar."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import pandas as pd
from openpyxl import Workbook

from flujos.armado.estructura import Estructura, analizar
from flujos.validacion.evaluador import ErrorExcel, Evaluador

TOLERANCIA = 0.05


@dataclass
class Control:
    nombre: str
    aprobado: bool
    detalle: str = ''
    nivel: str = 'ERROR'  # ERROR: no presentar sin revisar; AVISO: conviene mirarlo


def _saldo_final(ws: Any, ev: Evaluador, cabecera: int) -> float | None:
    for fila in ws.iter_rows(max_row=cabecera):
        for c in fila:
            if isinstance(c.value, str) and c.value.strip().upper().startswith(('SALDO FINAL', 'S.F')):
                for x in (ws.cell(c.row, c.column - 1) if c.column > 1 else None,
                          ws.cell(c.row, c.column + 1), ws.cell(c.row, c.column + 2)):
                    if x is not None and x.value is not None and not (isinstance(x.value, str) and not x.value.startswith('=')):
                        v = ev._valor_en(ws.title, x.row, x.column)
                        return float(v) if isinstance(v, (int, float)) else None
    return None


def saldos_contra_libro(wb: Workbook, ev: Evaluador, asignadas: dict[str, list[str]], cuentas: pd.DataFrame,
                        saldos_anteriores: dict[str, float | None] | None = None) -> list[Control]:
    controles = []
    for hoja, libros in asignadas.items():
        if hoja not in wb.sheetnames or len(libros) != 1:
            continue
        e = analizar(wb[hoja])
        fin = None if e is None else _saldo_final(wb[hoja], ev, e.fila_cabecera)
        fila = cuentas[cuentas.hoja == libros[0]]
        if fin is None or fila.empty:
            controles.append(Control(f'Saldo final {hoja}', False, 'No pude leer el saldo final de la hoja.', 'AVISO'))
            continue
        esperado = float(fila.iloc[0].saldo_final or 0)
        dif = round(fin - esperado, 2)
        detalle = f'Flujo {fin:,.2f} y libro {esperado:,.2f}: diferencia {dif:,.2f}.' if abs(dif) > TOLERANCIA else ''
        anterior = (saldos_anteriores or {}).get(hoja)
        if detalle and anterior is not None:
            dif_inicial = round(anterior - float(fila.iloc[0].saldo_inicial or 0), 2)
            if abs(dif_inicial - dif) <= TOLERANCIA:
                detalle += (f' Viene de arrastre: el saldo final del flujo anterior ya difería en {dif_inicial:,.2f} '
                            f'del saldo inicial del libro.')
        controles.append(Control(f'Saldo final {hoja}', abs(dif) <= TOLERANCIA, detalle))
    return controles


def secciones_contra_movimientos(wb: Workbook, asignadas: dict[str, list[str]], mov: pd.DataFrame,
                                 intereses: dict[str, float] | None = None) -> list[Control]:
    """Lo escrito en cada hoja debe sumar lo mismo que MOVIMIENTOS, por cuenta y tipo.

    Según la hoja, el interés overnight aparece además en DEBE o en HABER: esa es la única diferencia aceptada.
    """
    controles = []
    for hoja, libros in asignadas.items():
        if hoja not in wb.sheetnames:
            continue
        e = analizar(wb[hoja])
        if e is None:
            continue
        m = mov[mov.hoja.isin(libros)]
        debe = haber = 0.0
        for b in e.bloques:
            for f in range(b.primera, b.ultima + 1):
                d, h = wb[hoja].cell(f, e.col_debe).value, wb[hoja].cell(f, e.col_haber).value
                debe += d if isinstance(d, (int, float)) else 0
                haber += h if isinstance(h, (int, float)) else 0
        ing, egr = float(m[m.tipo == 'INGRESO'].monto.sum()), float(m[m.tipo == 'EGRESO'].monto.sum())
        interes = sum((intereses or {}).get(x, 0.0) for x in libros)
        aceptable = lambda d: abs(d) <= TOLERANCIA or abs(d - interes) <= TOLERANCIA  # noqa: E731, B023
        ok = aceptable(haber - egr) and aceptable(debe - ing)
        controles.append(Control(f'Movimientos {hoja}', ok, '' if ok else
                                 f'Egresos: hoja {haber:,.2f}, MOVIMIENTOS {egr:,.2f}. Ingresos: hoja {debe:,.2f}, '
                                 f'MOVIMIENTOS {ing:,.2f}.'))
    return controles


def subtotales_completos(wb: Workbook, estructuras: dict[str, Estructura]) -> list[Control]:
    """Cada subtotal debe abarcar todas las filas con monto de su sección."""
    malos = []
    for hoja in estructuras:
        e = analizar(wb[hoja])
        if e is None:
            continue
        for b in e.bloques:
            if b.fila_subtotal is None:
                continue
            formula = next((wb[hoja].cell(b.fila_subtotal, c).value for c in range(e.col_haber + 1, e.col_haber + 4)
                            if isinstance(wb[hoja].cell(b.fila_subtotal, c).value, str)), '')
            rangos = re.findall(r'SUM\(\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?\)', formula)
            if not rangos:
                continue
            desde = min(int(a) for a, _ in rangos)
            hasta = max(int(z or a) for a, z in rangos)
            filas = [f for f in range(b.primera, b.ultima + 1)
                     if isinstance(wb[hoja].cell(f, e.col_debe).value, (int, float))
                     or isinstance(wb[hoja].cell(f, e.col_haber).value, (int, float))]
            if filas and (min(filas) < desde or max(filas) > hasta):
                malos.append(f'{hoja} / {b.titulo} ({formula})')
    return [Control('Subtotales cubren sus filas', not malos, '; '.join(malos[:10]))]


def errores_de_formula(wb: Workbook, ev: Evaluador, previos: set[tuple[str, str]] | None = None) -> list[Control]:
    """Recalcula todo el libro. Separa los errores nuevos de los que ya venían en el flujo anterior."""
    previos = previos or set()
    nuevos, heredados = [], []
    for ws in wb.worksheets:
        for (f, c), celda in list(ws._cells.items()):
            if isinstance(celda.value, str) and celda.value.startswith('='):
                v = ev._valor_en(ws.title, f, c)
                if isinstance(v, ErrorExcel) or '#REF!' in celda.value:
                    (heredados if (ws.title, celda.coordinate) in previos else nuevos).append(
                        f'{ws.title}!{celda.coordinate}')
    salida = [Control('Sin errores de fórmula nuevos', not nuevos, ', '.join(nuevos[:15]))]
    if heredados:
        salida.append(Control('Errores que ya venían del mes anterior', False,
                              f'{len(heredados)}: ' + ', '.join(heredados[:12]), 'AVISO'))
    return salida


def errores_existentes(wb: Workbook) -> set[tuple[str, str]]:
    ev = Evaluador(wb)
    salida = set()
    for ws in wb.worksheets:
        for (f, c), celda in list(ws._cells.items()):
            if isinstance(celda.value, str) and celda.value.startswith('='):
                if '#REF!' in celda.value or isinstance(ev._valor_en(ws.title, f, c), ErrorExcel):
                    salida.add((ws.title, celda.coordinate))
    return salida


def resumen_cuadra(wb: Workbook, ev: Evaluador, mensual: str) -> list[Control]:
    """La hoja mensual y TOTAL deben llegar al mismo saldo final."""
    ws = wb[mensual]
    fila_fin = None
    for fila in ws.iter_rows(min_col=1, max_col=2):
        for c in fila:
            if isinstance(c.value, str) and re.match(r'\s*saldo\s+final', c.value, re.I):
                fila_fin = c.row
                break
        if fila_fin:
            break
    if fila_fin is None:
        return [Control('Hoja mensual cuadra', False, 'No encontré la fila de saldo final.', 'AVISO')]
    cab = None
    for f in range(1, 12):
        for c in ws[f]:
            if isinstance(c.value, str) and c.value.strip().upper() in ('TOTALES', 'TOTAL'):
                cab = c.column
    if cab is None:
        return [Control('Hoja mensual cuadra', False, 'No encontré la columna de totales.', 'AVISO')]
    v = ev._valor_en(mensual, fila_fin, cab)
    columnas = [ev._valor_en(mensual, fila_fin, c) for c in range(1, cab)]
    suma = sum(x for x in columnas if isinstance(x, (int, float)))
    ok = isinstance(v, (int, float)) and abs(v - suma) <= TOLERANCIA
    return [Control('Hoja mensual cuadra', ok, '' if ok else f'Saldo total {v} y suma por cuenta {suma:,.2f}.')]
