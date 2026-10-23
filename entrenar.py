"""Entrena (o reentrena) el modelo.

Uso: python entrenar.py historico.csv [modelo.pkl]
El histórico sale de extraer_historico.py (flujos cerrados) y crece con actualizar.py (meses revisados).
"""
import os, pickle, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
from modelo import Clasificador


def entrenar(h):
    """Devuelve (modelo, histórico usado)."""
    # si un mismo mes/empresa vino de dos fuentes (flujo cerrado y revisión), prima la revisión
    if 'fuente' in h:
        rev = set(zip(h[h.fuente == 'revision'].empresa, h[h.fuente == 'revision'].periodo))
        h = h[(h.fuente == 'revision') | ~pd.Series(list(zip(h.empresa, h.periodo))).isin(rev).values]
    return Clasificador().fit(h), h


if __name__ == '__main__':
    salida = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'modelo.pkl')
    m, h = entrenar(pd.read_csv(sys.argv[1]))
    pickle.dump(m, open(salida, 'wb'))
    print(f'Modelo entrenado con {len(h)} movimientos, {h.empresa.nunique()} empresas, '
          f'{h.periodo.nunique()} meses, {h.categoria.nunique()} categorías -> {salida}')
