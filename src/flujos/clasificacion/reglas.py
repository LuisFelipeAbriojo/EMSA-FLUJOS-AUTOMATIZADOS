"""Reglas de orden dentro de las secciones de transferencias entre cuentas."""
from __future__ import annotations

import re

from flujos.config import cargar
from flujos.lectura.normalizar import norm


def secciones_por_operacion() -> set[str]:
    return set(cargar()['transferencias']['secciones'])


def tipo_operacion(concepto: str) -> int:
    """Posición del tipo de operación (0 = primero del config ... len = otros) para ordenar transferencias."""
    tipos = cargar()['transferencias']['tipos']
    c = norm(concepto)
    return next((i for i, t in enumerate(tipos) if re.search(t['patron'], c)), len(tipos))
