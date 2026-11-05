"""Actualización del histórico con flujos corregidos. Todo con datos inventados."""
import datetime as dt
import os
import zipfile

import openpyxl
import pandas as pd
import pytest

from flujos.app import web
from flujos.historico.desde_flujos import abrir_zip, incorporar, validar_historico
from flujos.lectura.flujos import SIN_REVISAR, leer_flujo
from flujos.reportes.reporte_historico import escribir_reporte
from sinteticos import D, historico


def flujo(ruta, filas=None, autor=None):
    """Un RESUMEN FLUJO mínimo: una hoja de banco con secciones y fila de totales."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BCP$'
    ws.append(['BANCO DE PRUEBA DOLARES'])
    ws.append(['FECHA', 'CONCEPTO', 'DEBE', 'HABER', 'SUBTOTAL'])
    filas = filas if filas is not None else [
        ('DEPOSITOS', None, None, None),
        (D(2026, 9, 2), 'DEPOSITO CLIENTE ALFA', 100.0, None),
        (D(2026, 9, 3), 'DEPOSITO CLIENTE BETA', 50.0, None),
        ('PAGO PROVEEDORES', None, None, None),
        (D(2026, 9, 4), 'PAGO PROVEEDOR GAMA', None, 30.0),
        (None, None, 150.0, 30.0),
    ]
    for f in filas:
        ws.append(list(f))
    if autor:
        wb.properties.lastModifiedBy = autor
    wb.save(ruta)
    return str(ruta)


@pytest.fixture
def h():
    return historico(('ICP',))


def test_flujo_nuevo_se_agrega(tmp_path, h):
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')], h)
    assert len(res.historico) == len(h) + 3
    x = res.incorporados[0]
    assert (x.empresa, x.periodo, x.accion, x.antes, x.ahora) == ('ICP', '2026-09', 'Nuevo', 0, 3)
    assert (x.ingresos_usd, x.egresos_usd) == (150.0, 30.0)
    assert res.rechazados == []


def test_mismo_flujo_dos_veces_reemplaza_sin_duplicar(tmp_path, h):
    ruta = flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')
    una = incorporar([ruta], h)
    dos = incorporar([ruta], una.historico)
    assert len(dos.historico) == len(una.historico)
    assert dos.incorporados[0].accion == 'Reemplaza' and dos.incorporados[0].antes == 3


def test_reemplazar_no_toca_otras_empresas_ni_otros_meses(tmp_path):
    h = historico(('ICP', 'ONE'))
    periodo = sorted(h.periodo.astype(str).unique())[-1]
    anio, mes = periodo.split('-')
    nombre = {'07': 'JULIO', '08': 'AGOSTO', '09': 'SETIEMBRE'}[mes]
    filas = [('DEPOSITOS', None, None, None), (D(int(anio), int(mes), 2), 'DEPOSITO NUEVO', 10.0, None)]
    res = incorporar([flujo(tmp_path / f'RESUMEN FLUJO ICP {nombre} {anio}.xlsx', filas)], h)
    otros = h[(h.empresa != 'ICP') | (h.periodo.astype(str) != periodo)]
    assert len(res.historico) == len(otros) + 1
    pd.testing.assert_frame_equal(res.historico.iloc[:len(otros)][list(h.columns)].reset_index(drop=True),
                                  otros.reset_index(drop=True))


@pytest.mark.parametrize('nombre,texto', [
    ('ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx', 'clasificado'),
    ('ICP NOTAS SETIEMBRE 2026.xlsx', 'notas'),
    ('PROYECTADO ICP SETIEMBRE 2026.xlsx', 'proyectado'),
    ('RESUMEN FLUJO SETIEMBRE 2026.xlsx', 'empresa'),
    ('RESUMEN FLUJO ICP 2026.xlsx', 'mes'),
])
def test_rechaza_lo_que_no_es_un_flujo_entregado(tmp_path, h, nombre, texto):
    res = incorporar([flujo(tmp_path / nombre)], h)
    assert res.incorporados == [] and len(res.historico) == len(h)
    assert texto in res.rechazados[0][1].lower()


def test_rechaza_flujo_armado_que_nadie_volvio_a_guardar(tmp_path, h):
    ruta = flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', autor=SIN_REVISAR)
    assert leer_flujo(ruta).sin_revisar
    res = incorporar([ruta], h)
    assert res.incorporados == [] and 'nadie lo volvió a guardar' in res.rechazados[0][1]
    # al guardarlo en Excel cambia el autor y ya entra
    assert incorporar([flujo(ruta, autor='Quien revisa')], h).incorporados


def test_rechaza_nombre_de_un_mes_con_fechas_de_otro(tmp_path, h):
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP OCTUBRE 2026.xlsx')], h)
    assert res.incorporados == [] and 'otro mes' in res.rechazados[0][1]


def test_rechaza_dos_flujos_de_la_misma_empresa_y_mes(tmp_path, h):
    (tmp_path / 'a').mkdir()
    (tmp_path / 'b').mkdir()
    rutas = [flujo(tmp_path / x / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx') for x in 'ab']
    res = incorporar(rutas, h)
    assert res.incorporados == [] and len(res.rechazados) == 2


def test_rechaza_excel_danado_y_sigue_con_los_demas(tmp_path, h):
    malo = tmp_path / 'RESUMEN FLUJO ONE SETIEMBRE 2026.xlsx'
    malo.write_bytes(b'esto no es un excel')
    res = incorporar([str(malo), flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')], h)
    assert [x.empresa for x in res.incorporados] == ['ICP']
    assert 'No pude abrirlo' in res.rechazados[0][1]


@pytest.mark.parametrize('texto', ['11-Set-26', '11Set26', dt.date(2026, 9, 11)])
def test_fechas_escritas_a_mano_no_se_pierden(tmp_path, texto):
    filas = [('DEPOSITOS', None, None, None), (texto, 'DEPOSITO', 10.0, None)]
    r = leer_flujo(flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas))
    assert [m['fecha'] for m in r.movimientos] == ['2026-09-11']
    assert r.omitidas == [] and r.recuperadas == (0 if isinstance(texto, dt.date) else 1)


def test_fecha_como_numero_de_serie_de_excel(tmp_path):
    filas = [('DEPOSITOS', None, None, None), (46276, 'DEPOSITO', 10.0, None)]
    r = leer_flujo(flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas))
    assert [m['fecha'] for m in r.movimientos] == ['2026-09-11']


def test_fila_sin_fecha_se_avisa(tmp_path, h):
    filas = [('DEPOSITOS', None, None, None), (D(2026, 9, 2), 'DEPOSITO', 10.0, None), ('sin fecha', 'OTRO DEPOSITO', 5.0, None)]
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas)], h)
    assert res.incorporados[0].ahora == 1
    assert any('OTRO DEPOSITO' in d for _, d in res.avisos)


def test_una_suma_en_la_columna_de_concepto_no_es_un_movimiento(tmp_path):
    filas = [('DEPOSITOS', None, None, None), (D(2026, 9, 2), 'DEPOSITO', 10.0, None), (None, 10.0, 10.0, None)]
    assert leer_flujo(flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas)).omitidas == []


def test_avisa_si_lo_leido_no_suma_como_la_hoja(tmp_path, h):
    filas = [('DEPOSITOS', None, None, None), (D(2026, 9, 2), 'DEPOSITO', 10.0, None), (None, None, 15.0, 0.0)]
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas)], h)
    assert any('fila de totales' in d for _, d in res.avisos)


def test_avisa_secciones_nuevas_y_fechas_fuera_del_mes(tmp_path, h):
    filas = [('SECCION INVENTADA', None, None, None), (D(2026, 9, 2), 'UNO', 10.0, None), (D(2026, 9, 3), 'DOS', 10.0, None),
             (D(1926, 9, 4), 'TRES', 10.0, None)]
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx', filas)], h)
    textos = ' '.join(d for _, d in res.avisos)
    assert 'SECCION INVENTADA' in textos and '1926-09-04' in textos


def test_avisa_empresas_que_no_vinieron(tmp_path):
    res = incorporar([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')], historico(('ICP', 'ONE')))
    assert any('ONE' in d and 'No vino' in d for _, d in res.avisos)


def test_csv_que_no_es_historico_da_error_claro():
    with pytest.raises(ValueError, match='no parece un historico'):
        validar_historico(pd.DataFrame({'a': [1]}))


def test_zip_no_escribe_fuera_de_su_carpeta(tmp_path):
    ruta_flujo = flujo(tmp_path / 'origen.xlsx')
    z = tmp_path / 'flujos.zip'
    with zipfile.ZipFile(z, 'w') as f:
        f.write(ruta_flujo, '../../RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')
        f.writestr('carpeta/', '')
        f.writestr('leeme.txt', 'hola')
        f.writestr('~$temporal.xlsx', 'x')
    destino = tmp_path / 'x' / 'y'
    rutas, ignorados = abrir_zip(str(z), str(destino))
    assert len(rutas) == 1 and ignorados == ['leeme.txt']
    assert os.path.realpath(rutas[0]).startswith(os.path.realpath(destino))


def test_archivo_que_no_es_zip_da_error_claro(tmp_path):
    z = tmp_path / 'flujos.zip'
    z.write_bytes(b'no soy un zip')
    with pytest.raises(ValueError, match='no es un .zip'):
        abrir_zip(str(z), str(tmp_path / 'd'))


def test_app_entrega_historico_con_fecha_reporte_y_modelo(tmp_path, h, monkeypatch):
    monkeypatch.setattr(web, 'SALIDA', str(tmp_path / 'salida'))
    csv = tmp_path / 'historico.csv'
    h.to_csv(csv, index=False)
    z = tmp_path / 'flujos.zip'
    with zipfile.ZipFile(z, 'w') as f:
        f.write(flujo(tmp_path / 'a.xlsx'), 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')
    r = web.actualizar_historico([str(csv), str(z)], lambda _: None)
    assert (r['movimientos_antes'], r['movimientos'], r['incorporados']) == (len(h), len(h) + 3, 1)
    salidas = os.listdir(tmp_path / 'salida')
    assert 'modelo_web.pkl' in salidas
    assert any(x.startswith('historico_') for x in salidas) and any(x.startswith('REPORTE HISTORICO') for x in salidas)


def test_app_sin_flujos_validos_no_entrega_historico(tmp_path, h, monkeypatch):
    monkeypatch.setattr(web, 'SALIDA', str(tmp_path / 'salida'))
    csv = tmp_path / 'historico.csv'
    h.to_csv(csv, index=False)
    r = web.actualizar_historico([str(csv), flujo(tmp_path / 'ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx')], lambda _: None)
    assert r['incorporados'] == 0 and r['rechazados'] == 1
    assert [x for x in os.listdir(tmp_path / 'salida') if x.endswith(('.csv', '.pkl'))] == []


def test_app_pide_el_historico(tmp_path, monkeypatch):
    monkeypatch.setattr(web, 'SALIDA', str(tmp_path / 'salida'))
    with pytest.raises(ValueError, match='historico.csv'):
        web.actualizar_historico([flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')], lambda _: None)


def test_reporte_con_sus_tres_hojas(tmp_path, h):
    malo = tmp_path / 'ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx'
    res = incorporar([flujo(malo), flujo(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')], h)
    ruta = tmp_path / 'r.xlsx'
    escribir_reporte(str(ruta), res, '2026-11-05')
    wb = openpyxl.load_workbook(ruta)
    assert wb.sheetnames == ['Resumen', 'Rechazados', 'Avisos']
    assert wb['Resumen']['A6'].value == 'ICP' and wb['Rechazados']['A4'].value == malo.name
