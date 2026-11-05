import os

import openpyxl
import pandas as pd
import pytest

from flujos.clasificacion.clasificar import clasificar
from flujos.historico.actualizar import actualizar
from sinteticos import D, historico, hoja_libro, libro_basico

NOMBRE = 'ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx'


@pytest.fixture
def salida(tmp_path, modelo):
    libro = libro_basico(str(tmp_path / 'ICP SET 2026.xlsx'))
    hechos = clasificar([libro], 2026, 9, modelo, str(tmp_path / 'salida'))
    assert len(hechos) == 1
    return hechos[0], str(tmp_path / 'salida' / NOMBRE)


def test_resumen_del_resultado(salida):
    hecho, ruta = salida
    assert os.path.exists(ruta)
    assert (hecho['empresa'], hecho['movimientos'], hecho['cuentas'], hecho['descuadres']) == ('ICP', 9, 2, 0)


def test_hojas_y_columnas(salida):
    wb = openpyxl.load_workbook(salida[1])
    assert wb.sheetnames[0] == 'RESUMEN' and 'MOVIMIENTOS' in wb.sheetnames
    assert {'BBVA$', 'BCP S'} <= set(wb.sheetnames)
    assert wb['LISTAS'].sheet_state == 'hidden'
    mov = pd.read_excel(salida[1], sheet_name='MOVIMIENTOS')
    assert mov[mov.CONCEPTO == 'ITF'].CATEGORIA.tolist() == ['ITF']


def test_transferencias_ordenadas_por_tipo_y_luego_fecha(tmp_path, modelo):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    hoja_libro(wb, 'BBVA$', [(D(2026, 8, 31), 'INI', None, None),
                             (D(2026, 9, 1), 'VENTA DOLARES TC 3.40', None, 1.0),
                             (D(2026, 9, 2), 'BCR AL BCP', None, 2.0),
                             (D(2026, 9, 3), 'VENTA DOLARES TC 3.41', None, 3.0),
                             (D(2026, 9, 4), 'BCR AL BCP', None, 4.0)])
    libro = str(tmp_path / 'ICP SET 2026.xlsx')
    wb.save(libro)
    clasificar([libro], 2026, 9, modelo, str(tmp_path / 's'))
    ws = openpyxl.load_workbook(str(tmp_path / 's' / NOMBRE))['BBVA$']
    filas = [r for r in ws.iter_rows(min_row=8, values_only=True)]
    i = next(k for k, r in enumerate(filas) if r[0] == 'CARGO TRANSF. CTAS')
    conceptos = [r[1] for r in filas[i + 1:i + 5]]
    assert conceptos == ['BCR AL BCP', 'BCR AL BCP', 'VENTA DOLARES TC 3.40', 'VENTA DOLARES TC 3.41']


def test_empresa_desconocida_se_salta(tmp_path, modelo):
    libro = libro_basico(str(tmp_path / 'SIN CODIGO.xlsx'))
    assert clasificar([libro], 2026, 9, modelo, str(tmp_path / 's')) == []


def test_actualizar_aplica_correcciones_al_historico(salida, modelo):
    _, ruta = salida
    wb = openpyxl.load_workbook(ruta)
    ws = wb['MOVIMIENTOS']
    col = {c.value: c.column for c in ws[1]}
    for fila in range(2, ws.max_row + 1):
        if ws.cell(fila, col['CONCEPTO']).value == 'PAGO PROVEEDOR BETA':
            ws.cell(fila, col['CATEGORIA']).value = 'PAGO HABERES'
    wb.save(ruta)
    h = actualizar([ruta], historico(('ICP',)), modelo)
    nuevo = h[h.fuente == 'revision']
    assert set(nuevo.periodo) == {'2026-09'}
    assert nuevo[nuevo.concepto == 'PAGO PROVEEDOR BETA'].categoria.tolist() == ['PAGO HABERES']
    # la hoja por cuenta se rehízo con la corrección
    filas = list(openpyxl.load_workbook(ruta)['BBVA$'].iter_rows(values_only=True))
    i = next(k for k, r in enumerate(filas) if r[0] == 'PAGO HABERES')
    assert 'PAGO PROVEEDOR BETA' in [r[1] for r in filas[i + 1:i + 3]]


def test_actualizar_reemplaza_la_revision_anterior_del_mismo_mes(salida, modelo):
    _, ruta = salida
    primera = actualizar([ruta], historico(('ICP',)), modelo)
    segunda = actualizar([ruta], primera, modelo)
    assert len(segunda) == len(primera)


def test_fechas_sospechosas_sobreviven_a_actualizar(tmp_path, modelo):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    filas = [(D(2026, 8, 31), 'INI', None, None),
             *[(D(2026, 9, d), f'DEPOSITO CLIENTE {d}', 10.0, None) for d in range(1, 9)],
             (D(2026, 10, 8), 'PAGO PROVEEDOR BETA', None, 5.0),
             *[(D(2026, 9, d), f'DEPOSITO CLIENTE {d}', 10.0, None) for d in range(10, 20)]]
    hoja_libro(wb, 'BBVA$', filas)
    libro = str(tmp_path / 'ICP SET 2026.xlsx')
    wb.save(libro)
    clasificar([libro], 2026, 9, modelo, str(tmp_path / 's'))
    ruta = str(tmp_path / 's' / NOMBRE)

    def alertas():
        ws = openpyxl.load_workbook(ruta)['RESUMEN']
        return [c for c, in ws.iter_rows(max_col=1, values_only=True) if isinstance(c, str) and c.startswith('• ')]

    antes = alertas()
    assert len(antes) == 1
    actualizar([ruta], historico(('ICP',)), modelo)
    assert alertas() == antes
    mov = pd.read_excel(ruta, sheet_name='MOVIMIENTOS')
    assert mov[mov.CONCEPTO == 'PAGO PROVEEDOR BETA'].REVISAR.tolist() == ['SI']


def test_nombre_invalido_da_error_claro(tmp_path, modelo):
    ruta = tmp_path / 'cualquiera.xlsx'
    openpyxl.Workbook().save(ruta)
    with pytest.raises(ValueError, match='no parece un Excel clasificado'):
        actualizar([str(ruta)], historico(('ICP',)), modelo)
