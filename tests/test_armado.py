"""Armado del flujo sobre una plantilla inventada con la misma anatomía que los flujos reales."""
import datetime as dt

import openpyxl
import pandas as pd
import pytest
from openpyxl.styles import Font

from flujos.armado.armar import armar
from flujos.armado.estructura import Fila, analizar, filas_de_bloque
from flujos.armado.formulas import traducir
from flujos.armado.fuentes import asignar_hojas, nombre_salida
from flujos.armado.overnight import cruzan, emparejar
from flujos.armado.proceso import armar_flujo
from flujos.armado.rodar import nombre_hoja_mensual, rotular_mes, ubicar_tc
from flujos.clasificacion.clasificar import clasificar
from flujos.lectura.flujos import leer_flujo
from flujos.validacion.evaluador import ErrorExcel, Evaluador
from sinteticos import D, historico, libro_basico

TC_AGO, TC_SET = 3.30, 3.40


def plantilla(ruta, sumar_en_una_celda=False):
    """Flujo de agosto: TOTAL, hoja mensual, una cuenta en dólares, una en soles y Overn."""
    wb = openpyxl.Workbook()
    t = wb.active
    t.title = 'TOTAL'
    t['B2'] = 'EMPRESA DE PRUEBA'
    t['B4'] = 'Al 31 de Agosto del 2026'
    t['E4'], t['F4'], t['G4'] = 'JUL', 'AGO', 'SET'
    t['B9'] = 'DEPOSITOS'
    t['E9'] = 111.0
    t['F9'] = '=+AGOSTO!$F$8'
    t['B10'] = 'PROVEEDORES'
    t['E10'] = 222.0
    t['F10'] = '=+AGOSTO!$F$10'
    t['B15'] = 'SALDO FINAL'
    t['F15'] = '=SUM(F9:F10)'
    t['G15'] = '=SUM(G9:G10)'
    t['I13'] = 'AGO26/JUL26'
    t['I15'] = '=+F15/E15-1'
    t['E15'] = '=SUM(E9:E10)'

    m = wb.create_sheet('AGOSTO')
    m['B3'], m['E2'], m['F2'] = 'EMPRESA DE PRUEBA', 'AGO', 'JUL'
    m['E3'], m['F3'] = TC_AGO, 3.20
    m['B5'], m['C5'], m['D5'], m['E5'], m['F5'] = 'INGRESOS', 'BCP $', 'BCP S/.', 'OVERNIGHT', 'TOTALES'
    m['B8'] = 'DEPOSITOS'
    m['C8'] = "=+'BCP$'!E12"
    m['D8'] = "=+'BCP SOLES'!E10/E3"
    m['F8'] = '=SUM(C8:E8)'
    m['B9'] = 'TRANSF. ENTRE CTAS'
    m['C9'] = "=+'BCP$'!E20+99.5"  # ajuste a mano del mes
    m['E9'] = '=+Overn!D13'
    m['F9'] = '=SUM(C9:E9)'
    m['B10'] = 'PROVEEDORES'
    m['C10'] = "=+'BCP$'!E16"
    m['F10'] = '=SUM(C10:E10)'
    m['B11'] = 'INTERESES'
    m['C11'] = "=+'BCP$'!E24"
    m['F11'] = '=SUM(C11:E11)'
    m['B14'] = 'Saldo Final'

    def cuenta(nombre, filas_dep, filas_prov, cabecera, saldo):
        w = wb.create_sheet(nombre)
        w['A3'] = cabecera
        w['C5'], w['D5'] = 'SALDO INICIAL', saldo
        w['C6'] = 'SALDO FINAL'
        w['A7'], w['B7'], w['C7'], w['D7'] = 'FECHA', 'CONCEPTO', 'DEBE', 'HABER'
        return w

    b = cuenta('BCP$', 0, 0, 'CTA.CTE. M.E. N 000-000', 1000.0)
    b['A9'] = 'DEPOSITOS'
    b['A10'], b['B10'], b['C10'] = D(2026, 8, 3), 'DEPOSITO VIEJO 1', 10.0
    b['A11'], b['B11'], b['C11'] = D(2026, 8, 4), 'DEPOSITO VIEJO 2', 20.0
    b['A12'].font = Font(bold=True)
    b['E12'] = '=SUM(C10:C12)'
    b['A14'] = 'PAGO PROVEEDORES'
    b['A15'], b['B15'], b['D15'] = D(2026, 8, 5), 'PAGO VIEJO', 5.0
    b['E16'] = '=SUM(D15:D16)'
    b['A18'] = 'CARGO TRANSF. CTAS'
    b['E20'] = '=SUM(D19:D20)'
    b['A22'] = 'INTERESES OVERNIGHT'
    b['E24'] = '=SUM(C23:C24)'
    b['A26'] = 'ITF'
    b['E27'] = '=SUM(D27)' if sumar_en_una_celda else '=SUM(D27:D27)'
    b['C29'] = '=SUM(C8:C28)'
    b['D29'] = '=SUM(D8:D28)'
    b['D31'] = '=+D5+C29-D29'
    b['D6'] = '=+D31'

    s = cuenta('BCP SOLES', 0, 0, 'CTA.CTE. M.N. N 000-111', 660.0)
    s['A9'] = 'DEPOSITOS'
    s['E10'] = '=SUM(C10:C10)'
    s['C12'] = '=SUM(C8:C11)'
    s['D12'] = '=SUM(D8:D11)'
    s['D14'] = '=+D5+C12-D12'
    s['D6'] = '=+D14'

    o = wb.create_sheet('Overn')
    o['B4'], o['C4'], o['D4'], o['E4'] = 'FECHA', 'CONCEPTO', 'DEBE', 'HABER'
    o['A8'], o['A9'], o['A11'], o['A12'] = 'BCP $', 'BCP S/', 'BCP $', 'BCP S/'
    o['D13'] = '=SUM(D6:D12)'
    o['E13'] = '=SUM(E8:E12)'
    wb.save(ruta)
    # los valores guardados los calcula Excel en un archivo real; aquí los inyecta el evaluador
    from flujos.armado.guardar import guardar_con_valores
    guardar_con_valores(openpyxl.load_workbook(ruta), ruta)
    return str(ruta)


def _fuente(movs):
    def fuente(hoja, e, _vinculadas):
        return [movs.get(hoja, {}).get(b.categoria, []) for b in e.bloques]
    return fuente


@pytest.fixture
def resultado(tmp_path):
    movs = {'BCP$': {'DEPOSITOS': [Fila(D(2026, 9, d), f'DEPOSITO NUEVO {d}', 100.0) for d in (1, 2, 3, 4, 5)],
                     'PAGO PROVEEDORES': [Fila(D(2026, 9, 7), 'PAGO NUEVO', None, 40.0)]},
            'BCP SOLES': {'DEPOSITOS': [Fila(D(2026, 9, 8), 'DEPOSITO SOLES', 340.0)]}}
    return armar(plantilla(tmp_path / 'RESUMEN FLUJO ICP AGOSTO 2026.xlsx'), 'ICP', 2026, 9, TC_SET, _fuente(movs))


def test_renombra_la_hoja_mensual_y_sus_vinculos(resultado):
    wb = resultado.libro
    assert 'SETIEMBRE' in wb.sheetnames and 'AGOSTO' not in wb.sheetnames
    assert wb['TOTAL']['G9'].value == '=+SETIEMBRE!$F$8'


def test_congela_el_mes_anterior_y_crea_la_columna_nueva(resultado):
    t = resultado.libro['TOTAL']
    assert isinstance(t['F9'].value, float)  # agosto queda como valor
    assert t['G10'].value == '=+SETIEMBRE!$F$10'
    assert resultado.rodaje.columna_congelada == 'F' and resultado.rodaje.columna_nueva == 'G'


def test_rotulos_variacion_y_titulo(resultado):
    t = resultado.libro['TOTAL']
    assert t['I13'].value == 'SET26/AGO26'
    assert t['I15'].value == '=+G15/F15-1'
    assert t['B4'].value == 'Al 30 de Setiembre del 2026'


def test_tc_actual_pasa_a_anterior(resultado):
    m = resultado.libro['SETIEMBRE']
    assert (m['E3'].value, m['F3'].value) == (TC_SET, TC_AGO)
    assert (m['E2'].value, m['F2'].value) == ('SET', 'AGO')


def test_inserta_filas_y_mantiene_formulas(resultado):
    wb = resultado.libro
    w = wb['BCP$']
    e = analizar(w)
    dep = next(b for b in e.bloques if b.categoria == 'DEPOSITOS')
    assert len(filas_de_bloque(w, e, dep)) == 5
    assert w.cell(dep.ultima, 5).value == f'=SUM(C{dep.primera}:C{dep.ultima})'
    # el vínculo de la hoja mensual sigue al subtotal aunque se haya movido
    assert wb['SETIEMBRE']['C8'].value == f"=+'BCP$'!E{dep.ultima}"
    ev = Evaluador(wb)
    assert ev.valor('SETIEMBRE', 'C8') == 500.0
    assert ev.valor('SETIEMBRE', 'D8') == 100.0  # 340 / TC de setiembre
    assert ev.valor('SETIEMBRE', 'C10') == 40.0


def test_quita_filas_si_hay_menos_movimientos(tmp_path):
    movs = {'BCP$': {'DEPOSITOS': [Fila(D(2026, 9, 1), 'SOLO UNO', 7.0)]}}
    wb = armar(plantilla(tmp_path / 'p.xlsx'), 'ICP', 2026, 9, TC_SET, _fuente(movs)).libro
    assert Evaluador(wb).valor('SETIEMBRE', 'C8') == 7.0
    assert not any(isinstance(c.value, str) and 'VIEJO' in c.value for c in wb['BCP$']._cells.values())


def test_saldo_inicial_es_el_final_anterior_y_el_saldo_cierra(resultado):
    w = resultado.libro['BCP$']
    assert w['D5'].value == pytest.approx(1000.0 + 30.0 - 5.0)
    assert Evaluador(resultado.libro).valor('BCP$', 'D6') == pytest.approx(1025.0 + 500.0 - 40.0)


def test_no_arrastra_ajustes_a_mano_y_los_reporta(resultado):
    m = resultado.libro['SETIEMBRE']
    assert '99.5' not in m['C9'].value
    assert ('AGOSTO', 'C9', '+99.5') in resultado.rodaje.ajustes_quitados


def test_conserva_estilos_de_filas_existentes(tmp_path):
    ruta = plantilla(tmp_path / 'p.xlsx')
    wb = openpyxl.load_workbook(ruta)
    wb['BCP$']['B10'].font = Font(bold=True, color='FF0000')
    wb.save(ruta)
    movs = {'BCP$': {'DEPOSITOS': [Fila(D(2026, 9, d), 'NUEVO', 1.0) for d in (1, 2, 3, 4)]}}
    w = armar(ruta, 'ICP', 2026, 9, TC_SET, _fuente(movs)).libro['BCP$']
    for f in (10, 11, 12):
        assert w.cell(f, 2).font.bold


def test_suma_de_una_sola_celda_crece_con_la_seccion(tmp_path):
    ruta = plantilla(tmp_path / 'p.xlsx', sumar_en_una_celda=True)
    movs = {'BCP$': {'ITF': [Fila(D(2026, 9, d), 'ITF', None, 0.5) for d in (1, 2, 3)]}}
    wb = armar(ruta, 'ICP', 2026, 9, TC_SET, _fuente(movs)).libro
    e = analizar(wb['BCP$'])
    itf = next(b for b in e.bloques if b.categoria == 'ITF')
    assert Evaluador(wb)._valor_en('BCP$', itf.ultima, 5) == 1.5


def test_seccion_con_movimientos_sin_vinculo_crea_el_vinculo(tmp_path):
    ruta = plantilla(tmp_path / 'p.xlsx')
    wb = openpyxl.load_workbook(ruta)
    w = wb['BCP SOLES']
    w['A11'] = 'PAGO PROVEEDORES'  # la sección existe, pero ninguna línea de la hoja mensual la suma
    w['C12'], w['D12'] = None, None
    w['E12'] = '=SUM(D12:D12)'
    w['C14'], w['D14'] = '=SUM(C8:C13)', '=SUM(D8:D13)'
    w['D16'] = '=+D5+C14-D14'
    w['D6'] = '=+D16'
    wb.save(ruta)
    movs = {'BCP SOLES': {'PAGO PROVEEDORES': [Fila(D(2026, 9, 9), 'PAGO EN SOLES', None, 340.0)]},
            'BCP$': {'PAGO PROVEEDORES': [Fila(D(2026, 9, 9), 'PAGO', None, 1.0)]}}
    res = armar(ruta, 'ICP', 2026, 9, TC_SET, _fuente(movs))
    assert any('PAGO PROVEEDORES' in x for x in res.vinculos_nuevos)
    assert Evaluador(res.libro).valor('SETIEMBRE', 'D10') == pytest.approx(100.0)


def _mov_over(hoja, fecha, concepto, tipo, monto, moneda='USD'):
    return dict(hoja=hoja, banco='BCP', moneda=moneda, fecha=fecha, concepto=concepto, tipo=tipo, monto=monto)


def test_interes_con_cargo_del_mes_anterior():
    mov = pd.DataFrame([
        _mov_over('BCP$', dt.date(2026, 8, 31), 'OVERNIGHT DEL 31 AL 1 TASA 2.5%', 'EGRESO', 1000.0),
        _mov_over('BCP$', dt.date(2026, 9, 1), 'OVERNIGHT DEL 31 AL 1 TASA 2.5%', 'INGRESO', 1000.75),
        _mov_over('BCP$', dt.date(2026, 9, 2), 'OVERNIGHT DEL 2 AL 3 TASA 2.5%', 'EGRESO', 500.0),
        _mov_over('BCP$', dt.date(2026, 9, 3), 'OVERNIGTH DEL 2 AL 3 TEA 2.5', 'INGRESO', 500.1),
    ])
    overs = emparejar(mov)
    assert sorted(o.interes for o in overs) == [0.1, 0.75]


def test_mismo_texto_en_varios_meses_empareja_el_monto_correcto():
    t = 'OVERNIGHT DEL 2 AL 3 TASA 2.5%'
    mov = pd.DataFrame([_mov_over('B', dt.date(2026, 8, 2), t, 'EGRESO', 130000.0),
                        _mov_over('B', dt.date(2026, 8, 3), t, 'INGRESO', 130010.0),
                        _mov_over('B', dt.date(2026, 9, 2), t, 'EGRESO', 750000.0),
                        _mov_over('B', dt.date(2026, 9, 3), t, 'INGRESO', 750050.0)])
    assert sorted(o.interes for o in emparejar(mov)) == [10.0, 50.0]


def test_no_empareja_overnights_de_meses_distintos_con_los_mismos_dias():
    mov = pd.DataFrame([_mov_over('B', dt.date(2026, 8, 28), 'OVERNIGHT DEL 28 AL 29 TASA 2.5%', 'EGRESO', 100.0),
                        _mov_over('B', dt.date(2026, 8, 29), 'OVERNIGHT DEL 28 AL 29 TASA 2.5%', 'INGRESO', 100.5),
                        _mov_over('B', dt.date(2026, 9, 28), 'OVERNIGHT DEL 28 AL 29 TASA 2.7%', 'EGRESO', 90.0),
                        _mov_over('B', dt.date(2026, 9, 29), 'OVERNIGHT DEL 28 AL 29 TASA 2.7%', 'INGRESO', 90.2)])
    assert sorted(o.interes for o in emparejar(mov)) == [0.2, 0.5]


def test_over_solo_registra_los_que_cruzan_de_mes():
    mov = pd.DataFrame([
        _mov_over('B', dt.date(2026, 8, 31), 'OVERNIGHT DEL 31 AL 1', 'EGRESO', 100.0),
        _mov_over('B', dt.date(2026, 9, 1), 'OVERNIGHT DEL 31 AL 1', 'INGRESO', 100.1),
        _mov_over('B', dt.date(2026, 9, 10), 'OVERNIGHT DEL 10 AL 11', 'EGRESO', 200.0),
        _mov_over('B', dt.date(2026, 9, 11), 'OVERNIGHT DEL 10 AL 11', 'INGRESO', 200.1),
        _mov_over('B', dt.date(2026, 9, 30), 'OVERNIGHT DEL 30 AL 1', 'EGRESO', 300.0),
        _mov_over('B', dt.date(2026, 10, 1), 'OVERNIGHT DEL 30 AL 1', 'INGRESO', 300.1),
        _mov_over('B', dt.date(2026, 9, 12), 'OVERNIGHT DEL 12 AL 13', 'EGRESO', 400.0),  # abono perdido: no cruza
    ])
    salen, vuelven = cruzan(emparejar(mov), 2026, 9)
    assert [o.cargo for o in salen] == [300.0]
    assert [o.cargo for o in vuelven] == [100.0]


def test_empareja_cuentas_por_continuidad_de_saldo():
    cuentas = pd.DataFrame([dict(hoja='BCP S 009', moneda='PEN', saldo_inicial=500.0, movimientos=3),
                            dict(hoja='BCP S 089', moneda='PEN', saldo_inicial=700.0, movimientos=3),
                            dict(hoja='SBK$', moneda='USD', saldo_inicial=1.0, movimientos=1)])
    asignadas, avisos = asignar_hojas(cuentas, {'BCP-009': 500.0, 'BCP-089': 700.0, 'SCOT$': 9.0})
    assert asignadas == {'BCP-009': ['BCP S 009'], 'BCP-089': ['BCP S 089'], 'SCOT$': ['SBK$']}
    assert avisos == []


def test_cuenta_ambigua_avisa():
    cuentas = pd.DataFrame([dict(hoja='BCP S', moneda='PEN', saldo_inicial=1.0, movimientos=3)])
    asignadas, avisos = asignar_hojas(cuentas, {'BCP-009': 500.0, 'BCP-089': 700.0})
    assert asignadas == {} and 'no tiene una hoja clara' in avisos[0]


@pytest.mark.parametrize('actual,mes,nuevo', [('AGOSTO', 9, 'SETIEMBRE'), ('AGO', 9, 'SET'), ('AGO-26', 9, 'SET-26'),
                                              ('DIC', 1, 'ENE')])
def test_nombre_de_hoja_mensual(actual, mes, nuevo):
    assert nombre_hoja_mensual(actual, mes) == nuevo


@pytest.mark.parametrize('texto,esperado', [('TC AGOSTO', 'TC SETIEMBRE'), ('TC JULIO', 'TC AGOSTO'), ('AGO26/JUL26', 'SET26/AGO26'),
                                            ('Ago26 vs Jul26', 'Set26 vs Ago26'), ('Ago26 vs Ago25', 'Set26 vs Set25')])
def test_rotulos_de_mes(texto, esperado):
    assert rotular_mes(texto, 9) == esperado


def test_nombre_de_salida_conserva_el_patron():
    assert nombre_salida('x/RESUMEN FLUJO ESSA AGOSTO-2026.xlsx', 'AGOSTO', 'SETIEMBRE') == 'RESUMEN FLUJO ESSA SETIEMBRE-2026.xlsx'
    assert nombre_salida('FLUJO CAJA ONE AGOSTO 2026.xlsx', 'AGOSTO', 'SETIEMBRE') == 'FLUJO CAJA ONE SETIEMBRE 2026.xlsx'


def test_traduce_referencias():
    mapas = {'BCP$': lambda f, lado='una': f + 10}
    r = traducir("=+'BCP$'!E5+AGOSTO!G3/E3", 'AGOSTO', mapas, {'AGOSTO': 'SETIEMBRE'})
    assert r == "=+'BCP$'!E15+SETIEMBRE!G3/E3"
    assert traducir('=SUM(C5:C9)', 'BCP$', mapas, {}) == '=SUM(C15:C19)'
    assert traducir('=+#REF!', 'BCP$', mapas, {}) == '=+#REF!'


def test_evaluador():
    wb = openpyxl.Workbook()
    w = wb.active
    w.title = 'A B'
    w['A1'], w['A2'], w['A3'] = 10, 'texto', 5
    w['B1'] = '=SUM(A1:A3)'
    w['B2'] = "='A B'!A1*2+A3/5"
    w['B3'] = '=A1/0'
    w['B4'] = '=B3+1'
    w['B5'] = '=A9'
    w['B6'] = '=+#REF!'
    w['B7'] = '=A1%'
    ev = Evaluador(wb)
    assert ev.valor('A B', 'B1') == 15
    assert ev.valor('A B', 'B2') == 21
    assert ev.valor('A B', 'B3') == ErrorExcel('#DIV/0!')
    assert ev.valor('A B', 'B4') == ErrorExcel('#DIV/0!')
    assert ev.valor('A B', 'B5') == 0
    assert ev.valor('A B', 'B6') == ErrorExcel('#REF!')
    assert ev.valor('A B', 'B7') == 0.1


def test_tc_se_ubica_por_su_uso():
    wb = openpyxl.Workbook()
    w = wb.active
    w['H3'], w['H4'] = 3.3, 3.4  # el del mes es el que convierte los movimientos
    w['C10'] = "='BCP'!E1/H4"
    w['C11'] = "='BCP'!E2/H4"
    w['C12'] = "='BCP'!D5/H3"
    assert ubicar_tc(w) == ('H4', 'H3')


def test_proceso_completo_con_validaciones_y_notas(tmp_path, modelo):
    libro = libro_basico(str(tmp_path / 'ICP SET 2026.xlsx'))
    h = clasificar([libro], 2026, 9, modelo, str(tmp_path / 'c'))
    ruta_flujo = plantilla(tmp_path / 'RESUMEN FLUJO ICP AGOSTO 2026.xlsx')
    a = armar_flujo(h[0]['ruta'], ruta_flujo, 'ICP', 2026, 9, TC_SET, str(tmp_path / 's'))
    assert a.flujo.endswith('RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx')
    # el archivo se reabre con fórmulas y con valores
    f, v = openpyxl.load_workbook(a.flujo), openpyxl.load_workbook(a.flujo, data_only=True)
    assert str(f['BCP$']['D6'].value).startswith('=')
    assert isinstance(v['BCP$']['D6'].value, float)
    notas = openpyxl.load_workbook(a.notas)
    assert notas.sheetnames == ['Resumen', 'Validaciones', 'Alertas', 'Variaciones', 'Supuestos']
    assert notas['Resumen']['B4'].value == TC_SET
    # queda marcado hasta que una persona lo guarde: así no entra al histórico sin revisar
    assert leer_flujo(a.flujo).sin_revisar


def test_tc_invalido_da_error_claro(tmp_path):
    with pytest.raises(ValueError, match='tipo de cambio'):
        armar_flujo('x', 'y', 'ICP', 2026, 9, 34.37, str(tmp_path))


def test_no_modifica_la_plantilla(tmp_path):
    ruta = plantilla(tmp_path / 'p.xlsx')
    antes = open(ruta, 'rb').read()
    armar(ruta, 'ICP', 2026, 9, TC_SET, _fuente({}))
    assert open(ruta, 'rb').read() == antes
    assert not historico().empty
