"""Evaluador de las fórmulas que usan los flujos: aritmética, referencias entre hojas y SUM.

En el navegador no hay Excel ni LibreOffice. En los flujos reales la única función que aparece es SUM,
así que basta con este evaluador, que se verifica contra los valores guardados por Excel.
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Any

from openpyxl import Workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import range_boundaries
from openpyxl.utils.datetime import to_excel

Valor = float | str | bool | None


@dataclass(frozen=True)
class ErrorExcel:
    """Un error de Excel como valor: se propaga por las fórmulas igual que en la hoja."""

    codigo: str


class FuncionNoSoportada(ValueError):
    pass


_REF_HOJA = re.compile(r"^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$")


def separar_referencia(texto: str, hoja_actual: str) -> tuple[str, str]:
    """'BCP$'!E25 -> (BCP$, E25); E25 -> (hoja_actual, E25). Quita los $ de la referencia."""
    m = _REF_HOJA.match(texto)
    if not m:
        return hoja_actual, texto.replace('$', '')
    hoja = m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2)
    return hoja, m.group(3).replace('$', '')


class Evaluador:
    def __init__(self, wb: Workbook) -> None:
        self.wb = wb
        self._memo: dict[tuple[str, int, int], Any] = {}
        self._en_curso: set[tuple[str, int, int]] = set()
        self._limites: dict[str, tuple[int, int]] = {}

    def valor(self, hoja: str, celda: str) -> Any:
        """Valor calculado de una celda, como 'E25'."""
        min_c, min_f, _, _ = range_boundaries(celda)
        return self._valor_en(hoja, min_f, min_c)

    def _valor_en(self, hoja: str, fila: int, col: int) -> Any:
        clave = (hoja, fila, col)
        if clave in self._memo:
            return self._memo[clave]
        if hoja not in self.wb.sheetnames:
            return ErrorExcel('#REF!')
        celda = self.wb[hoja]._cells.get((fila, col))
        v = None if celda is None else celda.value
        if isinstance(v, str) and v.startswith('='):
            if clave in self._en_curso:
                return ErrorExcel('#CIRCULAR!')
            self._en_curso.add(clave)
            limite = sys.getrecursionlimit()
            sys.setrecursionlimit(max(limite, 20000))
            try:
                v = self._evaluar(v, hoja)
                if v is None:  # una referencia a una celda vacía vale 0
                    v = 0.0
            finally:
                sys.setrecursionlimit(limite)
                self._en_curso.discard(clave)
        elif isinstance(v, (dt.datetime, dt.date)):
            v = to_excel(v)
        elif isinstance(v, str) and re.fullmatch(r'#[A-Z/0!?]+[!?A]?', v):
            v = ErrorExcel(v)
        self._memo[clave] = v
        return v

    def _evaluar(self, formula: str, hoja: str) -> Any:
        toks = [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]
        return _Parser(toks, hoja, self).expresion()

    def rango(self, hoja: str, ref: str) -> Iterator[Any]:
        min_c, min_f, max_c, max_f = range_boundaries(ref)
        ws = self.wb[hoja] if hoja in self.wb.sheetnames else None
        if ws is None:
            yield ErrorExcel('#REF!')
            return
        if hoja not in self._limites:  # ws.max_row y ws.max_column recorren toda la hoja en cada llamada
            self._limites[hoja] = (max((f for f, _ in ws._cells), default=0), max((c for _, c in ws._cells), default=0))
        ultima_fila, ultima_col = self._limites[hoja]
        for f in range(min_f, min(max_f, ultima_fila) + 1):
            for c in range(min_c, min(max_c, ultima_col) + 1):
                yield self._valor_en(hoja, f, c)


def _numero(v: Any) -> float | ErrorExcel:
    if isinstance(v, ErrorExcel):
        return v
    if v is None or v == '':
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', ''))
    except ValueError:
        return ErrorExcel('#VALUE!')


class _Parser:
    """Descenso recursivo sobre los tokens de openpyxl, con la precedencia de Excel."""

    def __init__(self, toks: list[Token], hoja: str, ev: Evaluador) -> None:
        self.t, self.i, self.hoja, self.ev = toks, 0, hoja, ev

    def _actual(self) -> Token | None:
        return self.t[self.i] if self.i < len(self.t) else None

    def _binaria(self, ops: tuple[str, ...], siguiente: Any, f: Any) -> Any:
        izq = siguiente()
        while (a := self._actual()) is not None and a.type == Token.OP_IN and a.value in ops:
            self.i += 1
            der = siguiente()
            izq = f(a.value, izq, der)
        return izq

    def expresion(self) -> Any:
        return self._binaria(('=', '<>', '<', '>', '<=', '>='), self._concatena, _comparar)

    def _concatena(self) -> Any:
        return self._binaria(('&',), self._suma, lambda _o, a, b: f'{a if a is not None else ""}{b if b is not None else ""}')

    def _suma(self) -> Any:
        return self._binaria(('+', '-'), self._producto, _aritmetica)

    def _producto(self) -> Any:
        return self._binaria(('*', '/'), self._potencia, _aritmetica)

    def _potencia(self) -> Any:
        return self._binaria(('^',), self._unario, _aritmetica)

    def _unario(self) -> Any:
        a = self._actual()
        if a is not None and a.type == Token.OP_PRE:
            self.i += 1
            v = _numero(self._unario())
            if isinstance(v, ErrorExcel):
                return v
            return -v if a.value == '-' else v
        v = self._primario()
        while (a := self._actual()) is not None and a.type == Token.OP_POST:
            self.i += 1
            n = _numero(v)
            v = n if isinstance(n, ErrorExcel) else n / 100
        return v

    def _primario(self) -> Any:
        a = self._actual()
        if a is None:
            return None
        self.i += 1
        if a.type == Token.OPERAND:
            if a.subtype == Token.NUMBER:
                return float(a.value)
            if a.subtype == Token.TEXT:
                return a.value[1:-1].replace('""', '"')
            if a.subtype == Token.LOGICAL:
                return a.value.upper() == 'TRUE'
            if a.subtype == Token.ERROR:
                return ErrorExcel(a.value)
            hoja, ref = separar_referencia(a.value, self.hoja)
            if ':' in ref:
                return list(self.ev.rango(hoja, ref))
            if '#REF!' in ref or not re.fullmatch(r'[A-Z]{1,3}\d+', ref):
                return ErrorExcel('#REF!')
            min_c, min_f, _, _ = range_boundaries(ref)
            return self.ev._valor_en(hoja, min_f, min_c)
        if a.type == Token.PAREN and a.subtype == Token.OPEN:
            v = self.expresion()
            self.i += 1
            return v
        if a.type == Token.FUNC and a.subtype == Token.OPEN:
            nombre = a.value[:-1].upper()
            if nombre != 'SUM':
                raise FuncionNoSoportada(f'La fórmula usa {nombre}, que el evaluador aún no soporta.')
            args = []
            while (b := self._actual()) is not None and not (b.type == Token.FUNC and b.subtype == Token.CLOSE):
                if b.type == Token.SEP:
                    self.i += 1
                    continue
                args.append(self.expresion())
            self.i += 1
            return _sum(args)
        raise ValueError(f'No entiendo este elemento de la fórmula: {a.value}')


def _sum(args: list[Any]) -> float | ErrorExcel:
    total = 0.0
    for arg in args:
        for v in arg if isinstance(arg, list) else [arg]:
            if isinstance(v, ErrorExcel):
                return v
            # como en Excel: dentro de un rango los textos se ignoran
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                total += v
    return total


def _aritmetica(op: str, a: Any, b: Any) -> float | ErrorExcel:
    if isinstance(a, list) or isinstance(b, list):
        return ErrorExcel('#VALUE!')
    x, y = _numero(a), _numero(b)
    if isinstance(x, ErrorExcel):
        return x
    if isinstance(y, ErrorExcel):
        return y
    if op == '+':
        return x + y
    if op == '-':
        return x - y
    if op == '*':
        return x * y
    if op == '^':
        return x**y
    return ErrorExcel('#DIV/0!') if y == 0 else x / y


def _comparar(op: str, a: Any, b: Any) -> bool | ErrorExcel:
    if isinstance(a, ErrorExcel):
        return a
    if isinstance(b, ErrorExcel):
        return b
    x, y = (0.0 if a is None else a), (0.0 if b is None else b)
    if op == '=':
        return x == y
    if op == '<>':
        return x != y
    try:
        return {'<': x < y, '>': x > y, '<=': x <= y, '>=': x >= y}[op]
    except TypeError:
        return ErrorExcel('#VALUE!')
