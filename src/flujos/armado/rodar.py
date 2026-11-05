"""Pasa la hoja mensual y TOTAL de un mes al siguiente: nombre, TC, columna congelada y columna nueva."""
from __future__ import annotations

import calendar
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

from openpyxl import Workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from flujos.lectura.normalizar import norm

NOMBRES = ['', 'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE',
           'NOVIEMBRE', 'DICIEMBRE']
ABREV = ['', 'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SET', 'OCT', 'NOV', 'DIC']
_REF_COLUMNA = re.compile(r"(?<![A-Za-z0-9_'!$.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])")


def _es_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith('=')


def _como_el_original(original: str, nuevo: str) -> str:
    if original.isupper():
        return nuevo
    if original.istitle():
        return nuevo.title()
    return nuevo.lower()


def rotular_mes(texto: str, mes_nuevo: int) -> str:
    """Corre un mes los nombres de mes de un rótulo: 'TC AGOSTO'->'TC SETIEMBRE', 'AGO26/JUL26'->'SET26/AGO26'."""
    anterior = 12 if mes_nuevo == 1 else mes_nuevo - 1
    previo = 12 if anterior == 1 else anterior - 1
    trad = {NOMBRES[anterior]: NOMBRES[mes_nuevo], NOMBRES[previo]: NOMBRES[anterior],
            ABREV[anterior]: ABREV[mes_nuevo], ABREV[previo]: ABREV[anterior], 'SEPTIEMBRE': None, 'SEP': None}

    def cambia(m: re.Match) -> str:
        n = norm(m.group(1))
        destino = trad.get(n)
        if n == 'SEPTIEMBRE' and anterior == 9:
            destino = NOMBRES[mes_nuevo]
        if destino is None:
            return m.group(0)
        return _como_el_original(m.group(1), destino)

    return re.sub(r'(?<![A-Za-zÁÉÍÓÚ])([A-Za-zÁÉÍÓÚ]{3,10})(?=(\d{0,4})(?![A-Za-zÁÉÍÓÚ]))()', lambda m: cambia(m), texto)


def nombre_hoja_mensual(actual: str, mes: int) -> str:
    """AGOSTO->SETIEMBRE, AGO->SET, AGO-26->SET-26, conservando el estilo que usa la empresa."""
    m = re.fullmatch(r'([A-Za-zÁÉÍÓÚ]+)(.*)', actual)
    if not m:
        raise ValueError(f'No reconozco la hoja mensual "{actual}".')
    largo = len(m.group(1)) > 3
    return (NOMBRES[mes] if largo else ABREV[mes]) + m.group(2)


@dataclass
class Rodaje:
    hoja_mensual_anterior: str
    hoja_mensual_nueva: str
    columna_congelada: str | None = None
    columna_nueva: str | None = None
    celdas_congeladas: int = 0
    celdas_nuevas: int = 0
    tc_actual: str | None = None
    tc_anterior: str | None = None
    avisos: list[str] = field(default_factory=list)
    ajustes_quitados: list[tuple[str, str, str]] = field(default_factory=list)


def ubicar_tc(ws: Worksheet) -> tuple[str | None, str | None]:
    """Busca las dos celdas de TC de la hoja mensual: la más usada al convertir movimientos es la del mes."""
    candidatos = [c.coordinate for c in ws._cells.values()
                  if c.row <= 6 and isinstance(c.value, (int, float)) and 2.5 < c.value < 5]
    if not candidatos:
        return None, None
    usos = Counter()
    for c in ws._cells.values():
        if c.row >= 7 and _es_formula(c.value):
            for cand in candidatos:
                letra, fila = re.match(r'([A-Z]+)(\d+)', cand).groups()
                if re.search(rf'/\s*(?:[^/!]*!)?\$?{letra}\$?{fila}(?!\d)', c.value):
                    usos[cand] += 1
    orden = [c for c, _ in usos.most_common()] + [c for c in candidatos if c not in usos]
    return orden[0], (orden[1] if len(orden) > 1 else None)


def rodar_mensual(ws: Worksheet, valores: Worksheet, mes: int, tc: float, r: Rodaje) -> None:
    actual, anterior = ubicar_tc(ws)
    r.tc_actual, r.tc_anterior = actual, anterior
    if actual is None:
        r.avisos.append('No encontré las celdas de TC en la hoja mensual.')
        return
    if anterior is not None:
        ws[anterior].value = valores[actual].value
    quitar_ajustes(ws, r)
    ws[actual].value = tc
    # rótulos de TC arriba o al costado de esas celdas
    for c in ws._cells.values():
        if c.row <= 6 and isinstance(c.value, str) and not _es_formula(c.value) \
                and re.search(r'\bTC\b|T\.C|^[A-Za-z]{3}$', c.value):
            c.value = rotular_mes(c.value, mes)


_AJUSTE_FINAL = re.compile(r'((?:\s*[+\-]\s*\d+(?:\.\d+)?)+)$')


def quitar_ajustes(ws: Worksheet, r: Rodaje) -> None:
    """Los ajustes escritos a mano en un vínculo son del mes que termina: no se arrastran, se reportan."""
    for c in ws._cells.values():
        if _es_formula(c.value) and (('!' in c.value) or 'SUM(' in c.value):
            m = _AJUSTE_FINAL.search(c.value)
            if m:
                r.ajustes_quitados.append((ws.title, c.coordinate, m.group(1).replace(' ', '')))
                c.value = c.value[:m.start()]


def _links_por_columna(ws: Worksheet, hoja_mes: str) -> Counter:
    columnas: Counter = Counter()
    for c in ws._cells.values():
        if _es_formula(c.value) and hoja_mes in c.value:
            columnas[c.column] += 1
    return columnas


def _correr_columna(formula: str) -> str:
    """Pasa la fórmula a la columna siguiente: las referencias de la misma hoja se corren, las de otra hoja no."""
    tok = Tokenizer(formula)
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE and '!' not in t.value:
            t.value = re.sub(r'(?<![$A-Z])([A-Z]{1,3})(\$?\d+)',
                             lambda m: get_column_letter(column_index_from_string(m.group(1)) + 1) + m.group(2), t.value)
    return tok.render()


def _continuar_saldo(ws: Worksheet, col_x: int, r: Rodaje) -> None:
    """El saldo inicial del mes nuevo es el final del mes que termina, aunque la columna ya trajera otra fórmula."""
    previa, letra_x = get_column_letter(col_x - 1), get_column_letter(col_x)
    for c in [x for x in ws._cells.values() if x.column == col_x and _es_formula(x.value)]:
        m = re.fullmatch(rf'=\+?{previa}(\d+)', c.value)
        if not m:
            continue
        destino = ws.cell(c.row, col_x + 1)
        esperado = f'=+{letra_x}{m.group(1)}'
        if destino.value == esperado or destino.value == f'={letra_x}{m.group(1)}':
            continue
        if destino.value is not None:
            r.avisos.append(f'TOTAL!{destino.coordinate} traía {destino.value}: el saldo inicial ahora sigue al mes anterior.')
        destino.value = esperado


def rodar_total(ws: Worksheet, valores: Worksheet, anterior: str, anio: int, mes: int, r: Rodaje) -> None:
    """Congela la columna del mes anterior y crea los vínculos de la columna nueva."""
    columnas = _links_por_columna(ws, anterior)
    if not columnas:
        r.avisos.append('TOTAL no tiene vínculos a la hoja mensual.')
        return
    col_x = columnas.most_common(1)[0][0]
    col_n = col_x + 1
    r.columna_congelada, r.columna_nueva = get_column_letter(col_x), get_column_letter(col_n)
    letra_x, letra_n = r.columna_congelada, r.columna_nueva
    for c in [x for x in ws._cells.values() if x.column == col_x]:
        if _es_formula(c.value) and '!' in c.value:
            destino = ws.cell(c.row, col_n)
            if destino.value is None:
                destino.value = _correr_columna(c.value)
                r.celdas_nuevas += 1
            c.value = valores[c.coordinate].value
            r.celdas_congeladas += 1
    _continuar_saldo(ws, col_x, r)
    # variaciones del mes contra el anterior, rótulos, acumulado comparable y título
    for c in list(ws._cells.values()):
        v = c.value
        if _es_formula(v):
            comp = re.fullmatch(rf'=\+?{letra_x}(\d+)/{get_column_letter(col_x - 1)}(\d+)-1', v)
            if comp and c.column > col_n:
                c.value = f'=+{letra_n}{comp.group(1)}/{letra_x}{comp.group(2)}-1'
                continue
            acum = re.fullmatch(r'=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', v)
            if acum and acum.group(2) == acum.group(4) and c.row <= 20:
                a, b = column_index_from_string(acum.group(1)), column_index_from_string(acum.group(3))
                if b - a + 1 == mes - 1 and b < col_x - 1:
                    c.value = f'=SUM({acum.group(1)}{acum.group(2)}:{get_column_letter(b + 1)}{acum.group(4)})'
        elif isinstance(v, str) and c.row <= 20:
            texto = norm(v)
            if re.search(r'[A-Z]{3}\s*\d{0,2}\s*(/|VS)\s*[A-Z]{3}\s*\d\d', texto):
                c.value = rotular_mes(v, mes)
            t = re.fullmatch(r'AL\s+\d+\s+(?:DE\s+)?([A-Z]+)(\s+(?:DE\s+|DEL\s+)?)(\d{4})', texto)
            if t and int(t.group(3)) == anio and t.group(1) in NOMBRES:
                ultimo = calendar.monthrange(anio, mes)[1]
                c.value = _titulo(v, ultimo, NOMBRES[mes])


def _titulo(original: str, dia: int, mes: str) -> str:
    m = re.fullmatch(r'(\s*AL\s+)\d+(\s+(?:DE\s+)?)([A-Za-záéíóú]+)(.*)', original, re.I)
    if not m:
        return original
    return f'{m.group(1)}{dia}{m.group(2)}{_como_el_original(m.group(3), mes)}{m.group(4)}'


def correr_vinculos_a_total(ws: Worksheet, r: Rodaje) -> None:
    """Los chequeos de la hoja mensual contra la columna del mes en TOTAL pasan a la columna nueva."""
    if not r.columna_congelada:
        return
    patron = re.compile(rf'(TOTAL!\$?){r.columna_congelada}(\$?\d+)')
    for c in ws._cells.values():
        if _es_formula(c.value) and 'TOTAL!' in c.value:
            c.value = patron.sub(lambda m: f'{m.group(1)}{r.columna_nueva}{m.group(2)}', c.value)


def renombrar_mensual(wb: Workbook, actual: str, mes: int) -> str:
    nuevo = nombre_hoja_mensual(actual, mes)
    wb[actual].title = nuevo
    return nuevo
