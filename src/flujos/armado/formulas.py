"""Reescribe referencias de fórmulas cuando cambian filas o nombres de hoja."""
from __future__ import annotations

import re
from collections.abc import Callable

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

from flujos.validacion.evaluador import separar_referencia

MapaFilas = Callable[[int, str], int]  # (fila, 'ini'|'fin'|'una') -> fila nueva
_CELDA = re.compile(r'^(\$?)([A-Z]{1,3})(\$?)(\d+)$')


def _citar(hoja: str) -> str:
    if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*', hoja) and not re.fullmatch(r'[A-Za-z]{1,3}\d+', hoja):
        return hoja
    return "'" + hoja.replace("'", "''") + "'"


def _mover_celda(celda: str, mapa: MapaFilas | None, lado: str) -> str:
    m = _CELDA.match(celda)
    if not m or mapa is None:
        return celda
    return f'{m.group(1)}{m.group(2)}{m.group(3)}{mapa(int(m.group(4)), lado)}'


def traducir(formula: str, hoja_actual: str, mapas: dict[str, MapaFilas], renombres: dict[str, str]) -> str:
    """Devuelve la fórmula con filas y hojas actualizadas. No toca lo que no reconoce (como #REF!)."""
    tok = Tokenizer(formula)
    cambio = False
    for t in tok.items:
        if t.type != Token.OPERAND or t.subtype != Token.RANGE or '[' in t.value or '#REF!' in t.value:
            continue
        tiene_hoja = '!' in t.value
        hoja, _ = separar_referencia(t.value, hoja_actual)
        ref = t.value.rsplit('!', 1)[-1]
        mapa = mapas.get(hoja)
        if ':' in ref:
            a, b = ref.split(':', 1)
            nuevo = f"{_mover_celda(a, mapa, 'ini')}:{_mover_celda(b, mapa, 'fin')}"
        else:
            nuevo = _mover_celda(ref, mapa, 'una')
        if tiene_hoja:
            nuevo = f'{_citar(renombres.get(hoja, hoja))}!{nuevo}'
        if nuevo != t.value:
            t.value = nuevo
            cambio = True
    return tok.render() if cambio else formula
