"""Entrena el modelo a partir del histórico de movimientos ya categorizados."""
from __future__ import annotations

import pandas as pd

from flujos.clasificacion.modelo import Clasificador


def entrenar(h: pd.DataFrame) -> tuple[Clasificador, pd.DataFrame]:
    """Devuelve (modelo, histórico usado)."""
    # si un mismo mes/empresa vino de dos fuentes (flujo cerrado y revisión), prima la revisión
    if 'fuente' in h:
        rev = set(zip(h[h.fuente == 'revision'].empresa, h[h.fuente == 'revision'].periodo, strict=True))
        pares = pd.Series(list(zip(h.empresa, h.periodo, strict=True)))
        h = h[(h.fuente == 'revision') | ~pares.isin(rev).values]
    return Clasificador().fit(h), h
