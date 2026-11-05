import pickle
import sys
import types

import pandas as pd
import pytest

from flujos.clasificacion.entrenar import entrenar
from flujos.clasificacion.modelo import Clasificador, cargar_modelo, guardar_modelo
from sinteticos import historico


def _mov(concepto, tipo='EGRESO', empresa='SBSA', monto=100.0, itf=0, fila=10):
    return dict(empresa=empresa, hoja='BBVA$', banco='BBVA', moneda='USD', concepto=concepto, tipo=tipo, monto=monto,
                itf_misma_fila=itf, fila_origen=fila)


def test_memoria_repite_conceptos_conocidos(modelo):
    r = modelo.predict(pd.DataFrame([_mov('PAGO PROVEEDOR BETA'), _mov('DEPOSITO CLIENTE ALFA', 'INGRESO')]))
    assert list(r.categoria) == ['PAGO PROVEEDORES', 'DEPOSITOS']
    assert set(r.origen) == {'memoria'}


def test_generaliza_conceptos_nuevos(modelo):
    r = modelo.predict(pd.DataFrame([_mov('PAGO PROVEEDOR ZETA')]))
    assert r.categoria[0] == 'PAGO PROVEEDORES'
    assert r.origen[0] == 'modelo'


def test_solo_propone_categorias_del_tipo(modelo):
    r = modelo.predict(pd.DataFrame([_mov('XXXX', 'INGRESO')]))
    assert r.categoria[0] in {'DEPOSITOS', 'ABONO TRANSF. CTAS'}


@pytest.mark.parametrize('empresa', ['ICP', 'SBSA'])
def test_itf_de_la_misma_fila_siempre_va_a_itf(modelo, empresa):
    # SBSA tiene en su histórico un ITF de la misma fila dejado dentro del movimiento: no debe imitarlo
    r = modelo.predict(pd.DataFrame([_mov('DEPOSITO CLIENTE ALFA', 'INGRESO', empresa),
                                     _mov('ITF', 'EGRESO', empresa, monto=0.1, itf=1)]))
    assert r.categoria[1] == 'ITF'
    assert r.origen[1] == 'regla ITF'


def test_la_revision_prima_sobre_el_flujo_cerrado():
    h = historico(('ICP',))
    h['fuente'] = 'flujo cerrado'
    rev = h[h.concepto == 'PAGO PROVEEDOR BETA'].assign(fuente='revision', categoria='PAGO HABERES')
    _, usado = entrenar(pd.concat([h, rev]))
    assert set(usado[usado.empresa == 'ICP'].fuente) == {'revision'}


def test_guardar_y_cargar(tmp_path, modelo):
    ruta = str(tmp_path / 'm.pkl')
    guardar_modelo(modelo, ruta)
    assert cargar_modelo(ruta).predict(pd.DataFrame([_mov('PAGO PROVEEDOR BETA')])).categoria[0] == 'PAGO PROVEEDORES'


def test_abre_modelos_guardados_con_la_ruta_antigua(tmp_path):
    # antes de la reestructura la clase vivía en el módulo `modelo`
    viejo = types.ModuleType('modelo')
    clase = type('Clasificador', (), {'__module__': 'modelo'})
    viejo.Clasificador = clase
    sys.modules['modelo'] = viejo
    try:
        obj = clase()
        obj.n_entrenamiento = 123
        datos = pickle.dumps(obj)
    finally:
        del sys.modules['modelo']
    ruta = tmp_path / 'antiguo.pkl'
    ruta.write_bytes(datos)
    cargado = cargar_modelo(str(ruta))
    assert isinstance(cargado, Clasificador)
    assert cargado.n_entrenamiento == 123


def test_archivo_que_no_es_modelo_da_error_claro(tmp_path):
    ruta = tmp_path / 'malo.pkl'
    ruta.write_bytes(b'esto no es un modelo')
    with pytest.raises(ValueError, match='no es un modelo'):
        cargar_modelo(str(ruta))
