"""Overnights: empareja cargo y abono, calcula intereses y llena las hojas Over.

Reglas de Finanzas:
- El interés de un overnight es abono - cargo. El cargo puede estar en el mes anterior: el libro trae la historia.
- La hoja Over solo registra los overnights que cruzan de un mes a otro: el dinero no salió de la empresa,
  está unos días en el banco y vuelve con intereses.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import Any

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from flujos.config import cargar
from flujos.lectura.normalizar import es_overnight, norm

_HOJA_OVER = re.compile(r'^(OVER|OVERN|OVERNIGHT)$')
COL_FECHA, COL_CONCEPTO = 2, 3
DIAS_MAXIMOS_OVERNIGHT = 20  # un cargo y un abono más separados no son el mismo overnight


@dataclass
class Overnight:
    hoja: str
    banco: str
    moneda: str
    concepto: str
    fecha_cargo: dt.date
    cargo: float
    fecha_abono: dt.date | None
    abono: float | None

    @property
    def interes(self) -> float | None:
        return None if self.abono is None else round(self.abono - self.cargo, 2)


def _clave(concepto: Any) -> str:
    """'OVERNIGHT DEL 3 AL 4 TASA 2.5%' y 'OVERNIGTH DEL 3 AL 4 TEA 2.5' son el mismo overnight."""
    m = re.search(r'(?:DEL\s+)?(\d{1,2})\s+AL\s+(\d{1,2})', norm(concepto))
    return f'{int(m.group(1))}-{int(m.group(2))}' if m else norm(concepto)


def emparejar(mov: pd.DataFrame) -> list[Overnight]:
    """Empareja cada cargo con su abono: mismo texto en la misma cuenta, el abono posterior más cercano en monto."""
    if mov.empty:
        return []
    m = mov[mov.concepto.map(es_overnight)].copy()
    m['clave'] = m.concepto.map(_clave)
    salida = []
    for (hoja, _clave_over), g in m.groupby(['hoja', 'clave'], sort=False):
        cargos = g[g.tipo == 'EGRESO'].sort_values('fecha').to_dict('records')
        abonos = g[g.tipo == 'INGRESO'].sort_values('fecha').to_dict('records')
        for c in cargos:
            candidatos = [a for a in abonos if a['monto'] >= c['monto']
                          and 0 <= (a['fecha'] - c['fecha']).days <= DIAS_MAXIMOS_OVERNIGHT]
            a = min(candidatos, key=lambda x: ((x['fecha'] - c['fecha']).days, x['monto'] - c['monto']), default=None)
            if a is not None:
                abonos.remove(a)
            salida.append(Overnight(hoja, c['banco'], c['moneda'], c['concepto'], c['fecha'], c['monto'],
                                    a['fecha'] if a else None, a['monto'] if a else None))
    return salida


def cruzan(overs: list[Overnight], anio: int, mes: int) -> tuple[list[Overnight], list[Overnight]]:
    """Devuelve (los que salen este mes y vuelven después, los que salieron antes y vuelven este mes)."""
    ini = dt.date(anio, mes, 1)
    fin = dt.date(anio + (mes == 12), mes % 12 + 1, 1)
    salen = [o for o in overs if ini <= o.fecha_cargo < fin and (
        (o.fecha_abono is not None and o.fecha_abono >= fin) or (o.fecha_abono is None and _vence_el_mes_siguiente(o)))]
    vuelven = [o for o in overs if o.fecha_cargo < ini and o.fecha_abono is not None and ini <= o.fecha_abono < fin]
    return salen, vuelven


def _vence_el_mes_siguiente(o: Overnight) -> bool:
    """Sin abono todavía: 'OVERNIGHT DEL 30 AL 1' cruza de mes; 'DEL 10 AL 11' es un abono que no se encontró."""
    m = re.search(r'DEL\s+(\d{1,2})\s+AL\s+(\d{1,2})', norm(o.concepto))
    if m:
        return int(m.group(2)) < int(m.group(1)) or '/' in norm(o.concepto)[m.end():m.end() + 1]
    return o.fecha_cargo.day >= 25


def hoja_over(nombres: list[str]) -> str | None:
    return next((n for n in nombres if _HOJA_OVER.match(norm(n))), None)


def _coincide_rotulo(rotulo: Any, o: Overnight) -> bool:
    t = norm(rotulo) if rotulo else ''
    if not t or not t.startswith(o.banco[:3]):
        return False
    soles = any(k in t for k in ('S/', 'SOL', 'PEN'))
    return soles == (o.moneda == 'PEN')


def llenar_hoja_over(ws: Worksheet, valores: Worksheet, empresa: str, salen: list[Overnight], vuelven: list[Overnight],
                     hoja_mes: str, tc_actual: str, tc_anterior: str) -> list[str]:
    """Escribe los overnights que cruzan de mes en las filas que usa cada empresa. Devuelve los avisos."""
    disposicion = cargar()['over'].get(empresa)
    if disposicion is None:
        return [f'No conozco la disposición de la hoja {ws.title} de {empresa}.']
    avisos = []
    _rodar_saldo(ws, valores)
    filas_salen, filas_vuelven = disposicion['filas_salen'], disposicion['filas_vuelven']
    col_salen = 4 if disposicion['columna_salen'] == 'D' else 5
    col_vuelven = 5 if col_salen == 4 else 4
    for f in (*filas_salen, *filas_vuelven):
        for c in (COL_FECHA, COL_CONCEPTO, 4, 5):
            ws.cell(f, c).value = None
    hoja = "'" + hoja_mes.replace("'", "''") + "'" if not re.fullmatch(r'[A-Za-z]+', hoja_mes) else hoja_mes

    def escribir(lista: list[Overnight], filas: list[int], col: int, tc: str) -> None:
        libres = list(filas)
        for o in lista:
            f = next((x for x in libres if _coincide_rotulo(ws.cell(x, 1).value, o)), None)
            if f is None:
                f = next((x for x in libres if not ws.cell(x, 1).value), None)
                if f is None and libres:
                    f = libres[0]
                if f is None:
                    avisos.append(f'{ws.title}: no hay fila libre para el overnight {o.concepto}.')
                    continue
                ws.cell(f, 1).value = ws.cell(f, 1).value or f'{o.banco} {"$" if o.moneda == "USD" else "S/"}'
            libres.remove(f)
            ws.cell(f, COL_FECHA).value = dt.datetime.combine(o.fecha_cargo, dt.time())
            ws.cell(f, COL_CONCEPTO).value = o.concepto
            letra, numero = re.match(r'([A-Z]+)(\d+)', tc).groups()
            ws.cell(f, col).value = o.cargo if o.moneda == 'USD' else f'={o.cargo}/{hoja}!${letra}${numero}'

    escribir(salen, filas_salen, col_salen, tc_actual)
    # el que vuelve se convierte con el TC con que salió, para que netee contra el mes anterior
    escribir(vuelven, filas_vuelven, col_vuelven, tc_anterior)
    return avisos


def _rodar_saldo(ws: Worksheet, valores: Worksheet) -> None:
    """Saldo inicial = saldo final anterior, si la hoja lleva saldo (por ejemplo E3 = E2 + D16 - E16)."""
    for fila in ws.iter_rows(max_row=7):
        for c in fila:
            if isinstance(c.value, str) and c.value.startswith('='):
                m = re.match(r'=\+?([A-Z]+)(\d+)\+', c.value)
                if m and int(m.group(2)) < c.row and valores[c.coordinate].value is not None:
                    ws[f'{m.group(1)}{m.group(2)}'].value = valores[c.coordinate].value
                    return


