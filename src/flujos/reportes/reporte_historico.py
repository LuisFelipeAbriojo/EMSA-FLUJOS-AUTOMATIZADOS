"""Reporte de lo que entró al histórico: qué se agregó o reemplazó, qué se rechazó y qué conviene mirar."""
from __future__ import annotations

from typing import Any

from openpyxl import Workbook

from flujos.historico.desde_flujos import Resultado
from flujos.reportes.notas import AMARILLO, ROJO, _cabecera, _titulo


def _anchos(ws: Any, anchos: list[int]) -> None:
    for letra, ancho in zip('ABCDEFGHIJK', anchos, strict=False):
        ws.column_dimensions[letra].width = ancho


def escribir_reporte(ruta: str, res: Resultado, fecha: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Resumen'
    _titulo(ws, f'Actualización del histórico, {fecha}')
    ws['A2'] = f'Movimientos antes: {res.antes:,}. Ahora: {len(res.historico):,}.'
    ws['A3'] = f'Flujos incorporados: {len(res.incorporados)}. Rechazados: {len(res.rechazados)}. Avisos: {len(res.avisos)}.'
    _cabecera(ws, 5, ['Empresa', 'Periodo', 'Qué pasó', 'Movimientos antes', 'Movimientos ahora', 'Ingresos S/', 'Egresos S/',
                      'Ingresos US$', 'Egresos US$', 'Fechas en texto leídas', 'Archivo'])
    for i, x in enumerate(res.incorporados, 6):
        for j, v in enumerate((x.empresa, x.periodo, x.accion, x.antes, x.ahora, x.ingresos_pen, x.egresos_pen,
                               x.ingresos_usd, x.egresos_usd, x.recuperadas, x.archivo), 1):
            c = ws.cell(i, j, v)
            if isinstance(v, float):
                c.number_format = '#,##0.00'
    _anchos(ws, [10, 10, 11, 12, 12, 16, 16, 16, 16, 12, 46])

    ws = wb.create_sheet('Rechazados')
    _titulo(ws, 'Archivos que no entraron al histórico')
    _cabecera(ws, 3, ['Archivo', 'Motivo'])
    for i, (archivo, motivo) in enumerate(res.rechazados, 4):
        ws.cell(i, 1, archivo).fill = ROJO
        ws.cell(i, 2, motivo)
    _anchos(ws, [50, 95])

    ws = wb.create_sheet('Avisos')
    _titulo(ws, 'Lo que conviene mirar')
    _cabecera(ws, 3, ['Empresa y periodo', 'Detalle'])
    for i, (clave, detalle) in enumerate(res.avisos, 4):
        ws.cell(i, 1, clave).fill = AMARILLO
        ws.cell(i, 2, detalle)
    _anchos(ws, [18, 120])
    wb.save(ruta)
