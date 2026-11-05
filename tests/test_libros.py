import datetime as dt

import openpyxl
import pytest

from flujos.lectura.libros import leer_empresa, leer_fecha
from sinteticos import D, hoja_libro, libro_basico


def _libro(tmp_path, nombre='BBVA$', filas=(), **kw):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    hoja_libro(wb, nombre, list(filas), **kw)
    ruta = str(tmp_path / 'ICP PRUEBA.xlsx')
    wb.save(ruta)
    return ruta


def _diferencia(c):
    return round(c.saldo_inicial + c.ingresos - c.egresos - c.saldo_final, 2)


@pytest.mark.parametrize('texto,esperado', [
    ('1-Set-26', dt.date(2026, 9, 1)),
    ('26-Apr-29', dt.date(2029, 4, 26)),
    ('5 OCT 2026', dt.date(2026, 10, 5)),
    ('08/09/2026', dt.date(2026, 9, 8)),
    ('12.09.26', dt.date(2026, 9, 12)),
    ('8-Aug-26', dt.date(2026, 8, 8)),
])
def testleer_fechas_escritas_a_mano(texto, esperado):
    assert leer_fecha(texto) == esperado


@pytest.mark.parametrize('texto', ['TOTAL', '31-02-2026', '1-XYZ-26', ''])
def testleer_fechas_invalidas(texto):
    assert leer_fecha(texto) is None


def test_solo_toma_el_mes_y_cuadra(tmp_path):
    mov, cuentas = leer_empresa(libro_basico(str(tmp_path / 'ICP SET 2026.xlsx')), 2026, 9)
    assert set(mov.empresa) == {'ICP'}
    assert list(cuentas.hoja) == ['BBVA$', 'BCP S']  # SALDOS no es una cuenta
    for _, c in cuentas.iterrows():
        assert c.saldo_inicial == 1000.0
        assert _diferencia(c) == 0
    assert dict(zip(cuentas.hoja, cuentas.moneda, strict=True)) == {'BBVA$': 'USD', 'BCP S': 'PEN'}


def test_itf_en_la_misma_fila_se_separa(tmp_path):
    mov, _ = leer_empresa(libro_basico(str(tmp_path / 'ICP SET 2026.xlsx')), 2026, 9)
    itf = mov[mov.itf_misma_fila == 1]
    assert len(itf) == 1
    fila = itf.iloc[0]
    assert (fila.concepto, fila.tipo, fila.monto) == ('ITF', 'EGRESO', 0.10)
    assert fila.concepto_padre == 'DEPOSITO CLIENTE ALFA'
    # la otra parte de la fila sigue siendo el depósito completo
    padre = mov[(mov.fila_origen == fila.fila_origen) & (mov.itf_misma_fila == 0)].iloc[0]
    assert (padre.tipo, padre.monto) == ('INGRESO', 2000.0)


def test_dos_montos_grandes_en_la_misma_fila_no_son_itf(tmp_path):
    ruta = _libro(tmp_path, filas=[(D(2026, 8, 31), 'INI', None, None), (D(2026, 9, 1), 'CRUCE', 500.0, 400.0)])
    mov, _ = leer_empresa(ruta, 2026, 9, 'ICP')
    assert mov.itf_misma_fila.sum() == 0
    assert len(mov) == 2


def _agosto(desde, hasta):
    return [(D(2026, 8, d), f'DEP {d}', 100.0, None) for d in range(desde, hasta)]


def test_fila_con_dia_y_mes_invertidos_cuenta_por_ubicacion(tmp_path):
    filas = [(D(2026, 7, 31), 'INI', None, None), *_agosto(1, 9), (D(2026, 9, 8), 'INVERTIDA', 2000.0, 0.10),
             *_agosto(10, 20), *[(D(2026, 9, d), f'SET {d}', None, 50.0) for d in range(1, 10)]]
    ruta = _libro(tmp_path, filas=filas)
    ago, cu_ago = leer_empresa(ruta, 2026, 8, 'ICP')
    set_, cu_set = leer_empresa(ruta, 2026, 9, 'ICP')
    assert 'INVERTIDA' in set(ago.concepto)
    assert 'INVERTIDA' not in set(set_.concepto)
    assert _diferencia(cu_ago.iloc[0]) == 0
    assert _diferencia(cu_set.iloc[0]) == 0
    fila = ago[ago.concepto == 'INVERTIDA'].iloc[0]
    assert fila.fecha == dt.date(2026, 9, 8)  # conserva la fecha original y avisa
    assert 'Fecha en el libro' in fila.aviso
    assert 'fila' in cu_ago.iloc[0].alertas


def test_bloque_de_otro_mes_cuenta_por_ubicacion(tmp_path):
    filas = [(D(2026, 8, 31), 'INI', None, None), *[(D(2026, 9, d), f'A {d}', 10.0, None) for d in range(1, 21)],
             *[(D(2026, 8, 21), f'B {i}', 5.0, None) for i in range(17)],
             *[(D(2026, 9, d), f'C {d}', 10.0, None) for d in range(21, 30)]]
    mov, cu = leer_empresa(_libro(tmp_path, filas=filas), 2026, 9, 'ICP')
    assert len(mov) == 20 + 17 + 9
    assert _diferencia(cu.iloc[0]) == 0
    assert '17 movimientos' in cu.iloc[0].alertas


def test_fila_antigua_debajo_del_mes_no_altera_saldo_inicial(tmp_path):
    filas = [(D(2026, 8, 31), 'INI', None, None), (D(2026, 9, 1), 'X', 10.0, None), (D(2026, 9, 2), 'Y', 10.0, None),
             (D(2026, 3, 1), 'VIEJA', None, 3.0), (D(2026, 9, 3), 'Z', 10.0, None)]
    _, cu = leer_empresa(_libro(tmp_path, filas=filas, saldo=500.0), 2026, 9, 'ICP')
    assert cu.iloc[0].saldo_inicial == 500.0


def test_cuotas_futuras_no_se_tocan(tmp_path):
    filas = [(D(2026, 8, 31), 'INI', None, None), (D(2026, 9, 5), 'P', 10.0, None),
             *[(D(2026 + (m > 12), (m - 1) % 12 + 1, 15), f'CUOTA {m}', None, 1.0) for m in range(10, 16)]]
    mov, cu = leer_empresa(_libro(tmp_path, filas=filas), 2026, 9, 'ICP')
    assert list(mov.concepto) == ['P']
    assert cu.iloc[0].alertas == ''


@pytest.mark.parametrize('columna', ['DESCRIPCION DE TRANSACCIONES', 'CLIENTE'])
def test_libro_con_formato_elsa(tmp_path, columna):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BCP$'
    ws.append(['EMPRESA DE PRUEBA'])
    ws.append(['CTA.CTE. M.E. N 000-000'])
    ws.append(['CCC', 'FECHA', columna, 'COD_FCAJA', 'DEBE', 'HABER', 'SALDO', 'N OPERACION'])
    ws.append([None, D(2026, 8, 31), 'CIERRE AL 31/08', None, None, None, 1000.0])
    ws.append([None, '2026.09.01', 'DEPOSITO ALFA', 'DEPOSITOS', 100.0, None, 1100.0, 'X1'])
    ws.append([None, D(2026, 9, 2), 'PAGO BETA', 'PAGO PROVEEDORES', None, 40.0, 1060.0])
    ruta = str(tmp_path / 'ELSA SET 2026.xlsx')
    wb.save(ruta)
    mov, cu = leer_empresa(ruta, 2026, 9)
    assert list(mov.concepto) == ['DEPOSITO ALFA', 'PAGO BETA']
    assert list(mov.operacion)[0] == 'X1'
    assert _diferencia(cu.iloc[0]) == 0
    assert (cu.iloc[0].saldo_inicial, cu.iloc[0].saldo_final) == (1000.0, 1060.0)


def testleer_fecha_con_el_anio_primero():
    assert leer_fecha('2026.09.01') == dt.date(2026, 9, 1)
    assert leer_fecha('2026-09-01') == dt.date(2026, 9, 1)


def test_filas_sueltas_bajo_bloque_vacio_se_marcan(tmp_path):
    filas = [(D(2026, 8, 31), 'INI', None, None), (D(2026, 9, 1), 'NORMAL', 10.0, None),
             *[(None, None, None, None)] * 4,
             (D(2026, 9, 28), 'PROYECTADO', None, 5.0)]
    mov, cu = leer_empresa(_libro(tmp_path, filas=filas), 2026, 9, 'ICP')
    assert 'proyectado' in mov[mov.concepto == 'PROYECTADO'].iloc[0].aviso
    assert cu.iloc[0].saldo_final == 1010.0  # el saldo final no toma la fila suelta
