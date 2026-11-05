import importlib.util
import logging
from pathlib import Path

import openpyxl
import pytest

from flujos import config
from flujos.cli import main
from flujos.lectura.flujos import extraer
from sinteticos import D

RAIZ = Path(__file__).parent.parent


def _flujo_cerrado(ruta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BCP$'
    ws['C5'] = 'SALDO INICIAL'
    ws.append([])
    ws.append(['FECHA', 'CONCEPTO', 'DEBE', 'HABER'])
    ws.append(['DEPOSITOS'])
    ws.append([D(2026, 9, 1), 'DEPOSITO ALFA', 100.0, None])
    ws.append(['5-Set-26', 'DEPOSITO FECHA EN TEXTO', 50.0, None])
    ws.append([None, None, None, None, '=SUM(C8:C9)'])
    ws.append(['CARGO TRASNF CTAS'])
    ws.append([D(2026, 9, 2), 'BCR AL BCP', None, 30.0])
    wb.save(ruta)
    return str(ruta)


def test_extraer_toma_las_secciones_como_categoria(tmp_path, caplog):
    with caplog.at_level(logging.WARNING, logger='flujos'):
        filas = extraer(_flujo_cerrado(tmp_path / 'RESUMEN FLUJO ICP SETIEMBRE 2026.xlsx'))
    assert [(f['concepto'], f['categoria'], f['tipo']) for f in filas] == [
        ('DEPOSITO ALFA', 'DEPOSITOS', 'INGRESO'), ('BCR AL BCP', 'CARGO TRANSF. CTAS', 'EGRESO')]
    assert {f['periodo'] for f in filas} == {'2026-09'}
    assert '1 filas omitidas' in caplog.text


def test_extraer_ignora_proyectado(tmp_path):
    assert extraer(_flujo_cerrado(tmp_path / 'ICP PROYECTADO SET 2026.xlsx')) == []


def test_config_local_reemplaza_sin_perder_el_resto(tmp_path, monkeypatch):
    local = tmp_path / 'local.toml'
    local.write_text('[clasificacion]\numbral_revision = 0.75\n', encoding='utf-8')
    monkeypatch.setenv(config.VARIABLE_CONFIG_LOCAL, str(local))
    config.cargar.cache_clear()
    try:
        c = config.cargar()
        assert c['clasificacion']['umbral_revision'] == 0.75
        assert 'ICP' in c['empresas']['codigos']
    finally:
        monkeypatch.delenv(config.VARIABLE_CONFIG_LOCAL)
        config.cargar.cache_clear()


def test_config_local_inexistente_da_error_claro(monkeypatch):
    monkeypatch.setenv(config.VARIABLE_CONFIG_LOCAL, 'no-existe.toml')
    config.cargar.cache_clear()
    try:
        with pytest.raises(FileNotFoundError, match='configuración local'):
            config.cargar()
    finally:
        monkeypatch.delenv(config.VARIABLE_CONFIG_LOCAL)
        config.cargar.cache_clear()


def test_cli_modelo_inexistente_devuelve_error(tmp_path):
    assert main(['clasificar', str(tmp_path), '--modelo', str(tmp_path / 'nada.pkl')]) == 1


def _privacidad():
    spec = importlib.util.spec_from_file_location('verificar_privacidad', RAIZ / 'scripts' / 'verificar_privacidad.py')
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.mark.parametrize('ruta,contenido,bloquea', [
    ('datos.xlsx', b'x', True),
    ('carpeta/HISTORICO.CSV', b'x', True),
    ('modelo.pkl', b'x', True),
    ('copia.7z', b'x', True),
    ('grande.txt', b'a' * (600 * 1024), True),
    ('notas.md', b'cuenta 000' + b'-0000000-0-00', True),
    ('notas.md', b'RUC. 20' + b'000000000', True),
    ('modulo.py', b'version = "0.27.7"\numbral = 0.9\nfecha = "2026-09-30"\n', False),
])
def test_candado_de_privacidad(ruta, contenido, bloquea):
    assert bool(_privacidad().revisar(ruta, contenido)) is bloquea
