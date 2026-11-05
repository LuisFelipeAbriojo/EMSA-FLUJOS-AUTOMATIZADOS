import pytest

from flujos.clasificacion.reglas import secciones_por_operacion, tipo_operacion
from flujos.lectura.normalizar import banco, categoria, empresa_de, moneda, norm


def test_norm_quita_tildes_y_espacios():
    assert norm('  Depósito   añó ') == 'DEPOSITO ANO'


@pytest.mark.parametrize('nombre,esperado', [
    ('ICP SET 2026.xlsx', 'ICP'),
    ('RESUMEN FLUJO ESSA AGOSTO-2026.xlsx', 'ESSA'),
    ('FLUJO CAJA T1 OCTUBRE 2026.xlsx', 'T1'),
    ('OTRA COSA.xlsx', None),
])
def test_empresa_de(nombre, esperado):
    assert empresa_de(nombre) == esperado


@pytest.mark.parametrize('hoja,esperado', [
    ('BCP$', 'BCP'), ('SBK SOL', 'SCOT'), ('NAC DET', 'NACION'), ('BN', 'NACION'), ('F.Mutuo', 'FM'), ('Hoja1', 'OTRO'),
    ('BANBIFS', 'BIF'),
])
def test_banco(hoja, esperado):
    assert banco(hoja) == esperado


@pytest.mark.parametrize('hoja,cabecera,esperado', [
    ('BCP$', '', 'USD'), ('BCP S', '', 'PEN'), ('BBVA SOL', '', 'PEN'), ('BCP-009', '', 'PEN'),
    ('GIRO', 'CTA.CTE. M.E. N', 'USD'), ('GIRO', 'CTA.CTE. M.N. N', 'PEN'), ('PICH', '', 'USD'),
])
def test_moneda(hoja, cabecera, esperado):
    assert moneda(hoja, cabecera) == esperado


@pytest.mark.parametrize('seccion,esperado', [
    ('CARGO TRASNF CTAS', 'CARGO TRANSF. CTAS'), ('PRESTAMO TERCEROS', 'PRESTAMOS DE TERCEROS'), ('itf', 'ITF'),
])
def test_categoria_unifica_variantes(seccion, esperado):
    assert categoria(seccion) == esperado


def test_tipo_operacion_ordena_y_tolera_erratas():
    assert {tipo_operacion(c) for c in ('OVERNIGHT DEL 1 AL 2', 'OVERNIGTH DEL 1', 'OVENRIGHT 3')} == {0}
    assert tipo_operacion('BCR AL BCP') == tipo_operacion('TBK-TRF.CTAS PROPIAS BCR') == 1
    assert tipo_operacion('VENTA DOLARES TC 3.40') == 2
    assert tipo_operacion('RESCATE') == tipo_operacion('R.CPDADM123') == 3
    assert tipo_operacion('REEMBOLSO CAJA') == 4
    assert secciones_por_operacion() == {'ABONO TRANSF. CTAS', 'CARGO TRANSF. CTAS'}
