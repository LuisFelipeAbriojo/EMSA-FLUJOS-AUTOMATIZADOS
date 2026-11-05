"""Notas para la Jefa de Finanzas: un Excel aparte, sin tocar el formato del flujo."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string

from flujos.validacion.revision import Control

AZUL = PatternFill('solid', fgColor='DDEBF7')
ROJO = PatternFill('solid', fgColor='F8CBAD')
AMARILLO = PatternFill('solid', fgColor='FFF2CC')
VERDE = PatternFill('solid', fgColor='C6EFCE')
B = Font(bold=True)
UMBRAL_VARIACION = 0.30
MINIMO_VARIACION = 10000


@dataclass
class DatosNotas:
    empresa: str
    mes_nombre: str
    anio: int
    tc: float
    tc_anterior: Any
    controles: list[Control]
    por_revisar: pd.DataFrame
    alertas_fechas: list[str]
    sin_seccion: pd.DataFrame
    avisos: list[str]
    ajustes_quitados: list[tuple[str, str, str]]
    diferencia_cambio: list[tuple[str, float]]
    variaciones: list[tuple[str, float, float]] = field(default_factory=list)
    cuentas_sin_movimientos: list[str] = field(default_factory=list)
    vinculos_nuevos: list[str] = field(default_factory=list)


def variaciones_total(ws: Any, ev: Any, col_nueva: str | None) -> list[tuple[str, float, float]]:
    """Líneas de TOTAL que cambiaron mucho contra el mes anterior."""
    if not col_nueva:
        return []
    n = column_index_from_string(col_nueva)
    salida = []
    for f in sorted({f for f, _ in ws._cells}):
        rotulo = next((ws.cell(f, c).value for c in range(1, n) if isinstance(ws.cell(f, c).value, str)
                       and not ws.cell(f, c).value.startswith('=')), None)
        a, p = ev._valor_en(ws.title, f, n), ev._valor_en(ws.title, f, n - 1)
        if rotulo and isinstance(a, (int, float)) and isinstance(p, (int, float)) and max(abs(a), abs(p)) >= MINIMO_VARIACION:
            cambio = (a - p) / abs(p) if p else 1.0
            if abs(cambio) >= UMBRAL_VARIACION:
                salida.append((' '.join(str(rotulo).split()), a, p))
    return salida


def _titulo(ws: Any, texto: str) -> None:
    ws['A1'] = texto
    ws['A1'].font = Font(bold=True, size=13)


def _cabecera(ws: Any, fila: int, nombres: list[str]) -> None:
    for i, n in enumerate(nombres, 1):
        c = ws.cell(fila, i, n)
        c.font, c.fill = B, AZUL


def escribir_notas(ruta: str, d: DatosNotas) -> dict[str, int]:
    """Escribe las notas y devuelve el conteo de errores y avisos."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Resumen'
    errores = [c for c in d.controles if not c.aprobado and c.nivel == 'ERROR']
    avisos = [c for c in d.controles if not c.aprobado and c.nivel == 'AVISO']
    _titulo(ws, f'{d.empresa} – Notas del flujo de {d.mes_nombre} {d.anio}')
    filas = [
        ('Resultado', 'LISTO PARA REVISAR' if not errores else f'{len(errores)} CONTROLES SIN PASAR'),
        ('TC de fin de mes', d.tc), ('TC del mes anterior', d.tc_anterior),
        ('Controles aprobados', sum(c.aprobado for c in d.controles)), ('Errores', len(errores)), ('Avisos', len(avisos)),
        ('Filas que siguieron por revisar', len(d.por_revisar)), ('Fechas sospechosas', len(d.alertas_fechas)),
        ('Movimientos sin sección en el flujo', len(d.sin_seccion)),
        ('Ajustes a mano del mes anterior retirados', len(d.ajustes_quitados)),
    ]
    for i, (a, b) in enumerate(filas, 3):
        ws.cell(i, 1, a).font = B
        ws.cell(i, 2, b)
    ws['B3'].fill = VERDE if not errores else ROJO
    ws.column_dimensions['A'].width = 44
    ws.column_dimensions['B'].width = 26

    w = wb.create_sheet('Validaciones')
    _titulo(w, 'Controles automáticos')
    _cabecera(w, 3, ['CONTROL', 'RESULTADO', 'DETALLE'])
    for i, c in enumerate(d.controles, 4):
        w.cell(i, 1, c.nombre)
        r = w.cell(i, 2, 'OK' if c.aprobado else c.nivel)
        r.fill = VERDE if c.aprobado else ROJO if c.nivel == 'ERROR' else AMARILLO
        w.cell(i, 3, c.detalle)
    w.column_dimensions['A'].width = 40
    w.column_dimensions['B'].width = 11
    w.column_dimensions['C'].width = 110

    w = wb.create_sheet('Alertas')
    _titulo(w, 'Puntos a mirar')
    _cabecera(w, 3, ['TIPO', 'DETALLE'])
    i = 4
    grupos = [('Fecha sospechosa', d.alertas_fechas), ('Armado', d.avisos),
              ('Vínculo nuevo en la hoja mensual', d.vinculos_nuevos),
              ('Cuenta sin movimientos', d.cuentas_sin_movimientos),
              ('Ajuste a mano retirado', [f'{h}!{c}: {m} (era del mes anterior)' for h, c, m in d.ajustes_quitados])]
    for tipo, lista in grupos:
        for x in lista:
            w.cell(i, 1, tipo)
            w.cell(i, 2, x)
            i += 1
    for r in d.sin_seccion.to_dict('records'):
        w.cell(i, 1, 'Sin sección en el flujo').fill = ROJO
        w.cell(i, 2, f"{r.get('hoja')}: {r.get('categoria')} – {r.get('fecha')} – {r.get('concepto')} – {r.get('monto')}")
        i += 1
    for r in d.por_revisar.to_dict('records'):
        w.cell(i, 1, 'Siguió por revisar').fill = AMARILLO
        w.cell(i, 2, f"{r.get('hoja')}: {r.get('categoria')} – {r.get('fecha')} – {r.get('concepto')} – {r.get('monto')}")
        i += 1
    w.column_dimensions['A'].width = 32
    w.column_dimensions['B'].width = 120

    w = wb.create_sheet('Variaciones')
    _titulo(w, f'Líneas de TOTAL que cambiaron {UMBRAL_VARIACION:.0%} o más contra el mes anterior')
    _cabecera(w, 3, ['LÍNEA', 'ESTE MES', 'MES ANTERIOR', 'CAMBIO'])
    for i, (r, a, p) in enumerate(d.variaciones, 4):
        w.cell(i, 1, r)
        w.cell(i, 2, a).number_format = '#,##0.00'
        w.cell(i, 3, p).number_format = '#,##0.00'
        w.cell(i, 4, (a - p) / abs(p) if p else None).number_format = '0%'
    w.column_dimensions['A'].width = 40
    for col in 'BCD':
        w.column_dimensions[col].width = 16

    w = wb.create_sheet('Supuestos')
    _titulo(w, 'Reglas aplicadas en este armado')
    reglas = [
        f'TC de fin de mes: {d.tc}. Es el que ingresó quien armó el flujo y es único para todas las empresas.',
        'Interés overnight = abono - cargo. El cargo puede ser del mes anterior.',
        'La hoja Over registra solo los overnights que cruzan de un mes a otro.',
        'El ITF que viene en la misma fila del movimiento va a la sección ITF.',
        'En ABONO / CARGO TRANSF. CTAS: por tipo de operación y luego por fecha.',
        'Una fila con un mes distinto al de sus vecinas se cuenta según su ubicación en el libro.',
        'Los ajustes escritos a mano en los vínculos del mes anterior no se arrastran.',
    ]
    for i, t in enumerate(reglas, 3):
        w.cell(i, 1, t)
    if d.diferencia_cambio:
        f = 4 + len(reglas)
        w.cell(f, 1, 'Diferencia de cambio estimada por cuenta (saldo convertido menos saldo de la hoja). '
                     'No se aplicó: es una sugerencia hasta confirmar la regla.').font = B
        _cabecera(w, f + 1, ['CUENTA', 'DIFERENCIA'])
        for k, (c, v) in enumerate(d.diferencia_cambio, f + 2):
            w.cell(k, 1, c)
            w.cell(k, 2, v).number_format = '#,##0.00'
        w.cell(f + 2 + len(d.diferencia_cambio), 1, 'Total').font = B
        w.cell(f + 2 + len(d.diferencia_cambio), 2, sum(v for _, v in d.diferencia_cambio)).number_format = '#,##0.00'
    w.column_dimensions['A'].width = 100
    w.column_dimensions['B'].width = 16
    wb.save(ruta)
    return dict(errores=len(errores), avisos=len(avisos))
