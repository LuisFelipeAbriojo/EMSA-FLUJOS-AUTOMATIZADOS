"""Arma el flujo del mes sobre una copia del flujo del mes anterior."""
from __future__ import annotations

import logging
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import get_column_letter

from flujos.armado.estructura import Estructura, Fila, analizar
from flujos.armado.formulas import MapaFilas, traducir
from flujos.armado.overnight import Overnight, hoja_over, llenar_hoja_over
from flujos.armado.reacomodo import reemplazar_bloques
from flujos.armado.rodar import (
    _AJUSTE_FINAL,
    Rodaje,
    correr_vinculos_a_total,
    nombre_hoja_mensual,
    renombrar_mensual,
    rodar_mensual,
    rodar_total,
)
from flujos.validacion.evaluador import separar_referencia

log = logging.getLogger(__name__)

# (hoja, estructura, filas de esa hoja con vínculo a la mensual) -> movimientos de cada bloque
Fuente = Callable[[str, Estructura, set[int]], list[list[Fila]] | None]


@dataclass
class Resultado:
    libro: Any
    rodaje: Rodaje
    estructuras: dict[str, Estructura] = field(default_factory=dict)
    hojas_sin_estructura: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    vinculadas: dict[str, set[int]] = field(default_factory=dict)
    vinculos_nuevos: list[str] = field(default_factory=list)


def filas_vinculadas(wb: Any, mensual: str) -> dict[str, set[int]]:
    """Por hoja, las filas a las que apuntan la hoja mensual y TOTAL."""
    salida: dict[str, set[int]] = {}
    for hoja in (mensual, 'TOTAL'):
        for celda in wb[hoja]._cells.values():
            if isinstance(celda.value, str) and celda.value.startswith('='):
                for t in Tokenizer(celda.value).items:
                    if t.type == Token.OPERAND and t.subtype == Token.RANGE and '!' in t.value:
                        h, ref = separar_referencia(t.value, hoja)
                        m = re.fullmatch(r'[A-Z]+(\d+)', ref)
                        if m:
                            salida.setdefault(h, set()).add(int(m.group(1)))
    return salida


def completar_vinculos(wb: Any, mensual: str, res: Resultado, hoja: str, e: Estructura,
                       contenido: list[list[Fila]]) -> list[str]:
    """Una sección con movimientos y sin vínculo perdería ese dinero en la hoja mensual: se crea el vínculo.

    La fila sale de la misma sección en otra cuenta, y la columna de los demás vínculos de esta hoja.
    """
    ws = wb[mensual]
    nuevos = []
    propias = res.vinculadas.get(hoja, set())
    cita = "'" + hoja.replace("'", "''") + "'"
    enlaces = [c for c in ws._cells.values() if isinstance(c.value, str) and c.value.startswith('=')
               and (f'{cita}!' in c.value or re.search(rf'(?<![\w\'.]){re.escape(hoja)}!', c.value))]
    if not enlaces:
        return nuevos
    columna = max({c.column for c in enlaces}, key=lambda x: sum(1 for c in enlaces if c.column == x))
    conversion = next((m.group(0) for c in enlaces if c.column == columna
                       for m in [re.search(r'/\s*[^/]+$', c.value)] if m), '')
    for b, filas in zip(e.bloques, contenido, strict=True):
        if not filas or b.ultima in propias or b.fila_subtotal is None:
            continue
        destino = None
        for otra, est in res.estructuras.items():
            if otra == hoja or destino:
                continue
            for x in est.bloques:
                if x.categoria == b.categoria and x.tipo == b.tipo and x.ultima in res.vinculadas.get(otra, set()):
                    cita_otra = "'" + otra.replace("'", "''") + "'"
                    patron = rf'(?:{re.escape(cita_otra)}|(?<![\w\'.]){re.escape(otra)})!\$?[A-Z]+\$?{x.ultima}(?!\d)'
                    celda = next((c for c in ws._cells.values() if isinstance(c.value, str) and c.value.startswith('=')
                                  and re.search(patron, c.value)), None)
                    if celda is not None:
                        destino = celda.row
                        break
        if destino is None:
            res.avisos.append(f'{hoja}: la sección {b.titulo} tiene movimientos y no sé a qué línea de la hoja mensual va.')
            continue
        celda = ws.cell(destino, columna)
        letra = re.sub(r'\d', '', celda.coordinate)
        col_sub = next((c for c in range(e.col_haber + 1, e.col_haber + 4)
                        if isinstance(wb[hoja].cell(b.fila_subtotal, c).value, str)), e.col_haber + 1)
        suma = f"{cita}!{get_column_letter(col_sub)}{b.ultima}"
        if celda.value is None:
            celda.value = f'=+{suma}{conversion}'
        elif isinstance(celda.value, str) and celda.value.startswith('='):
            celda.value = f'{celda.value}+{suma}{conversion}'
        else:
            continue
        nuevos.append(f'{letra}{destino} de la hoja mensual ahora suma {hoja} / {b.titulo}.')
    return nuevos


def hoja_mensual(wb: Any) -> str:
    nombres = wb.sheetnames
    if 'TOTAL' not in nombres or nombres.index('TOTAL') + 1 >= len(nombres):
        raise ValueError('El flujo anterior no tiene una hoja TOTAL seguida de la hoja mensual.')
    return nombres[nombres.index('TOTAL') + 1]


def _rodar_saldo_inicial(ws: Any, valores: Any, cabecera: int) -> bool:
    """Saldo inicial del mes = saldo final del mes anterior, como valor."""
    celdas = {}
    for fila in ws.iter_rows(max_row=cabecera):
        for c in fila:
            if isinstance(c.value, str) and not c.value.startswith('='):
                t = c.value.strip().upper().replace('.', '')
                if t.startswith(('SALDO INICIAL', 'SI')) and len(t) <= 14:
                    celdas['ini'] = _al_costado(ws, c)
                if t.startswith(('SALDO FINAL', 'SF')) and len(t) <= 14:
                    celdas['fin'] = _al_costado(ws, c)
    if 'ini' in celdas and 'fin' in celdas and celdas['ini'] and celdas['fin']:
        ws[celdas['ini']].value = valores[celdas['fin']].value
        return True
    return False


def _control_saldo(ws: Any, cabecera: int, saldo: float | None) -> None:
    """Algunas hojas comparan su saldo con el del libro escrito a mano (F6 = D6 - E6): se pone el del mes."""
    for c in list(ws._cells.values()):
        if c.row <= cabecera and isinstance(c.value, str):
            m = re.fullmatch(r'=\+?([A-Z]+\d+)-([A-Z]+\d+)', c.value)
            if m:
                otra = ws[m.group(2)]
                if isinstance(otra.value, (int, float)) and not isinstance(otra.value, bool):
                    otra.value = saldo


def _quitar_ajustes_de_hoja(ws: Any, r: Rodaje) -> None:
    for c in ws._cells.values():
        if isinstance(c.value, str) and c.value.startswith('=') and 'SUM(' in c.value:
            m = _AJUSTE_FINAL.search(c.value)
            if m:
                r.ajustes_quitados.append((ws.title, c.coordinate, m.group(1).replace(' ', '')))
                c.value = c.value[:m.start()]


def _al_costado(ws: Any, c: Any) -> str | None:
    # normalmente a la derecha del rótulo; en algunas hojas el saldo está a su izquierda
    izquierda = ws.cell(c.row, c.column - 1) if c.column > 1 else None
    if izquierda is not None and izquierda.value is not None:
        if not isinstance(izquierda.value, str | bool) or str(izquierda.value).startswith('='):
            return izquierda.coordinate
    for salto in (1, 2):
        x = ws.cell(c.row, c.column + salto)
        if x.value is not None:
            return x.coordinate
    return None


def _vinculos_a_movimientos(wb: Any, mensual: str, estructuras: dict[str, Estructura],
                            mapas: dict[str, MapaFilas]) -> list[str]:
    """Una fórmula que apunta a una fila de movimiento (no a un subtotal) no se puede arrastrar sola: se avisa."""
    avisos = []
    interiores = {h: {f for b in e.bloques for f in range(b.primera, b.ultima)} for h, e in estructuras.items() if h in mapas}
    for hoja in ('TOTAL', mensual):
        for celda in wb[hoja]._cells.values():
            if not (isinstance(celda.value, str) and celda.value.startswith('=')):
                continue
            for t in Tokenizer(celda.value).items:
                if t.type == Token.OPERAND and t.subtype == Token.RANGE and '!' in t.value and ':' not in t.value:
                    h, ref = separar_referencia(t.value, hoja)
                    m = re.fullmatch(r'[A-Z]+(\d+)', ref)
                    if m and int(m.group(1)) in interiores.get(h, ()):
                        avisos.append(f'{hoja}!{celda.coordinate} apunta a una fila de movimiento ({t.value}): revisar.')
    return avisos


def armar(plantilla: str, empresa: str, anio: int, mes: int, tc: float, fuente: Fuente,
          overs: tuple[list[Overnight], list[Overnight]] | None = None,
          saldo_libro: dict[str, float] | None = None) -> Resultado:
    """Devuelve el libro del mes. No guarda nada: la plantilla original no se toca.

    `fuente(hoja, estructura)` da los movimientos de cada bloque, o None para dejar la hoja como está.
    """
    wb = openpyxl.load_workbook(plantilla)
    valores = openpyxl.load_workbook(plantilla, data_only=True)
    anterior = hoja_mensual(wb)
    nuevo = nombre_hoja_mensual(anterior, mes)
    r = Rodaje(anterior, nuevo)
    res = Resultado(wb, r)

    mapas: dict[str, MapaFilas] = {}
    over = hoja_over(wb.sheetnames)
    res.vinculadas = filas_vinculadas(wb, anterior)
    for nombre in wb.sheetnames:
        if nombre in ('TOTAL', anterior, 'Resumen', over):
            continue
        ws = wb[nombre]
        e = analizar(ws)
        if e is None:
            res.hojas_sin_estructura.append(nombre)
            continue
        res.estructuras[nombre] = e
        if not _rodar_saldo_inicial(ws, valores[nombre], e.fila_cabecera):
            res.avisos.append(f'{nombre}: no encontré las celdas de saldo inicial y final.')
        _control_saldo(ws, e.fila_cabecera, (saldo_libro or {}).get(nombre))
        _quitar_ajustes_de_hoja(ws, r)
        contenido = fuente(nombre, e, res.vinculadas.get(nombre, set()))
        if contenido is None:
            continue
        res.vinculos_nuevos += completar_vinculos(wb, anterior, res, nombre, e, contenido)
        mapas[nombre] = reemplazar_bloques(ws, e, contenido)

    rodar_total(wb['TOTAL'], valores['TOTAL'], anterior, anio, mes, r)
    rodar_mensual(wb[anterior], valores[anterior], mes, tc, r)
    correr_vinculos_a_total(wb[anterior], r)
    renombrar_mensual(wb, anterior, mes)
    if over is not None and overs is not None and r.tc_actual:
        res.avisos += llenar_hoja_over(wb[over], valores[over], empresa, overs[0], overs[1], nuevo, r.tc_actual,
                                       r.tc_anterior or r.tc_actual)

    res.avisos += _vinculos_a_movimientos(wb, anterior if nuevo not in wb.sheetnames else nuevo, res.estructuras, mapas)

    renombres = {anterior: nuevo}
    for ws in wb.worksheets:
        for celda in list(ws._cells.values()):
            v = celda.value
            if isinstance(v, str) and v.startswith('='):
                # las referencias a la propia hoja también cambian de fila
                celda.value = traducir(v, anterior if ws.title == nuevo else ws.title, mapas, renombres)
    wb.calculation.fullCalcOnLoad = True
    return res
