"""Configuración: la general viene con el paquete y un archivo local puede agregar o reemplazar valores."""
from __future__ import annotations

import os
import tomllib
from functools import lru_cache
from importlib import resources
from typing import Any

VARIABLE_CONFIG_LOCAL = 'FLUJOS_CONFIG'


def _fusionar(base: dict[str, Any], extra: dict[str, Any]) -> dict[str, Any]:
    salida = dict(base)
    for clave, valor in extra.items():
        if isinstance(valor, dict) and isinstance(salida.get(clave), dict):
            salida[clave] = _fusionar(salida[clave], valor)
        else:
            salida[clave] = valor
    return salida


@lru_cache(maxsize=1)
def cargar() -> dict[str, Any]:
    """Devuelve la configuración. La ruta del archivo local sale de la variable FLUJOS_CONFIG."""
    texto = resources.files('flujos.recursos').joinpath('config.toml').read_text(encoding='utf-8')
    config = tomllib.loads(texto)
    ruta_local = os.environ.get(VARIABLE_CONFIG_LOCAL)
    if ruta_local:
        if not os.path.exists(ruta_local):
            raise FileNotFoundError(f'No encuentro el archivo de configuración local: {ruta_local}')
        with open(ruta_local, 'rb') as f:
            config = _fusionar(config, tomllib.load(f))
    return config
