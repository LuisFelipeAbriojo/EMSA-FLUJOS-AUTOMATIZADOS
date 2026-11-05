"""Pruebas de punta a punta con datos reales. Se saltan si no existe la carpeta privada (variable FLUJOS_DATOS)."""
import glob
import os

import openpyxl
import pandas as pd
import pytest

from flujos.clasificacion.clasificar import clasificar
from flujos.clasificacion.entrenar import entrenar

DATOS = os.environ.get('FLUJOS_DATOS', '')
pytestmark = pytest.mark.skipif(not os.path.isdir(DATOS), reason='sin carpeta de datos privados (FLUJOS_DATOS)')


@pytest.fixture(scope='module')
def resultado(tmp_path_factory):
    h = pd.read_csv(sorted(glob.glob(os.path.join(DATOS, 'historico', 'historico_*.csv')))[0])
    salida = str(tmp_path_factory.mktemp('salida'))
    modelo, _ = entrenar(h)
    libros = sorted(glob.glob(os.path.join(DATOS, 'Libros', '*.xlsx')))
    return clasificar(libros, 2026, 9, modelo, salida), salida


def test_setiembre_cuadra_en_todas_las_cuentas(resultado):
    hechos, _ = resultado
    assert len(hechos) >= 8
    assert {h['empresa']: h['descuadres'] for h in hechos if h['descuadres']} == {}


def test_igual_a_la_linea_base(resultado):
    """Salvo la hora de generación, cada celda coincide con la salida anterior a la reestructura."""
    _, salida = resultado
    base = os.path.join(DATOS, 'baseline_hito1', 'salida')
    if not os.path.isdir(base):
        pytest.skip('sin línea base')
    for ruta in glob.glob(os.path.join(base, '*.xlsx')):
        a = openpyxl.load_workbook(ruta)
        b = openpyxl.load_workbook(os.path.join(salida, os.path.basename(ruta)))
        assert [x for x in b.sheetnames if x != 'OVERNIGHTS'] == a.sheetnames
        for s in a.sheetnames:
            for fa, fb in zip(a[s].iter_rows(values_only=True), b[s].iter_rows(values_only=True), strict=True):
                if s == 'RESUMEN' and isinstance(fa[0], str) and fa[0].startswith('Generado'):
                    continue
                assert fa == fb, os.path.basename(ruta)
