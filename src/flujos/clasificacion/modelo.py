"""Clasificador local de movimientos bancarios.

Dos capas:
  1. Memoria: si el mismo concepto (misma empresa, mismo tipo) ya se clasificó siempre igual, se repite.
  2. Modelo lineal (regresión logística) sobre el texto del concepto + empresa, banco, moneda, tipo y monto.
Solo se proponen categorías que esa empresa ya usó para ese tipo (ingreso/egreso).
"""
from __future__ import annotations

import os
import pickle
import re
import unicodedata
from collections import Counter, defaultdict
from typing import Any

import numpy as np
import pandas as pd
from scipy.sparse import hstack
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import OneHotEncoder


def limpiar(s: Any) -> str:
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().upper()
    s = re.sub(r'\d+', '#', s)  # números de operación, facturas, fechas
    return re.sub(r'\s+', ' ', s).strip()


def _prep(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    d['txt'] = d.concepto.map(limpiar)
    d['raw'] = d.concepto.map(lambda x: re.sub(r'\s+', ' ', str(x).upper()).strip())
    d['cod'] = d.raw.map(lambda x: ' '.join('C' + c for c in re.findall(r'\d{6,}', x)) or 'SINCOD')
    d['rango'] = pd.cut(np.log10(d.monto.astype(float).clip(lower=0.01)), [-9, 0, 1, 2, 3, 4, 5, 6, 99],
                        labels=False).astype(str)
    return d


class Clasificador:
    CAT = ['empresa', 'tipo', 'banco', 'moneda', 'rango']

    def fit(self, d: pd.DataFrame) -> Clasificador:
        d = _prep(d)
        # el ITF que viene en la misma fila siempre va a ITF (aunque algún flujo cerrado lo dejara en el movimiento)
        if 'itf_misma_fila' in d:
            d.loc[d.itf_misma_fila.fillna(0).astype(int) == 1, 'categoria'] = 'ITF'
        self.tf_c = TfidfVectorizer(analyzer='char_wb', ngram_range=(2, 5), min_df=2, sublinear_tf=True)
        self.tf_w = TfidfVectorizer(analyzer='word', ngram_range=(1, 2), sublinear_tf=True)
        self.tf_k = TfidfVectorizer(analyzer='word', token_pattern=r'\S+')
        self.oh = OneHotEncoder(handle_unknown='ignore')
        X = hstack([self.tf_c.fit_transform(d.txt), self.tf_w.fit_transform(d.txt), self.tf_k.fit_transform(d.cod),
                    self.oh.fit_transform(d[self.CAT]) * 1.5]).tocsr()
        self.lr = LogisticRegression(C=10, max_iter=3000, class_weight='balanced').fit(X, d.categoria)
        def top(v: Counter) -> tuple[str, float, int]:
            return v.most_common(1)[0][0], v.most_common(1)[0][1] / sum(v.values()), sum(v.values())

        m1, m2 = defaultdict(Counter), defaultdict(Counter)
        for r, t, e, ti, c in zip(d.raw, d.txt, d.empresa, d.tipo, d.categoria, strict=True):
            m1[(r, e, ti)][c] += 1
            m2[(t, e, ti)][c] += 1
        self.mem_exacta = {k: top(v) for k, v in m1.items()}
        self.mem_texto = {k: top(v) for k, v in m2.items()}
        # categorías permitidas por empresa y tipo, y su orden habitual en las hojas
        self.permitidas = d.groupby(['empresa', 'tipo']).categoria.agg(lambda s: set(s)).to_dict()
        self.permitidas_tipo = d.groupby('tipo').categoria.agg(lambda s: set(s)).to_dict()
        if 'orden' in d:
            self.orden = d.groupby(['empresa', 'tipo', 'categoria']).orden.median().to_dict()
        self.n_entrenamiento = len(d)
        return self

    def orden_de(self, empresa: str, tipo: str, cat: str) -> float:
        o = getattr(self, 'orden', {})
        return o.get((empresa, tipo, cat), 999)

    def predict(self, d: pd.DataFrame) -> pd.DataFrame:
        """Devuelve DataFrame con categoria, confianza, alternativa, origen."""
        d = _prep(d.reset_index(drop=True))
        X = hstack([self.tf_c.transform(d.txt), self.tf_w.transform(d.txt), self.tf_k.transform(d.cod),
                    self.oh.transform(d[self.CAT]) * 1.5]).tocsr()
        P = self.lr.predict_proba(X)
        clases = self.lr.classes_
        out = []
        for i, (e, ti, r, t) in enumerate(zip(d.empresa, d.tipo, d.raw, d.txt, strict=True)):
            ok = self.permitidas.get((e, ti)) or self.permitidas_tipo.get(ti, set(clases))
            p = np.where(np.isin(clases, list(ok)), P[i], 0)
            p = p / p.sum() if p.sum() > 0 else P[i]
            j = np.argsort(-p)
            cat, conf, alt, origen = clases[j[0]], float(p[j[0]]), clases[j[1]], 'modelo'
            m = self.mem_exacta.get((r, e, ti)) or self.mem_texto.get((t, e, ti))
            if m and m[1] >= 0.9:
                if m[0] != cat:
                    alt = cat
                cat, origen = m[0], 'memoria'
                conf = max(conf, m[1]) if m[2] >= 2 else max(conf, 0.85)
            out.append((cat, round(conf, 3), alt, origen))
        res = pd.DataFrame(out, columns=['categoria', 'confianza', 'alternativa', 'origen'])
        if 'itf_misma_fila' in d:  # ITF de la misma fila (BBVA y otros): siempre a la categoría ITF
            i = np.where(d.itf_misma_fila.fillna(0).astype(int) == 1)[0]
            res.loc[i, ['categoria', 'confianza', 'origen']] = ['ITF', 0.99, 'regla ITF']
        return res


class _CargadorCompatible(pickle.Unpickler):
    """Permite abrir modelos guardados antes de la reestructura, cuando la clase vivía en el módulo `modelo`."""

    def find_class(self, module: str, name: str) -> Any:
        return super().find_class('flujos.clasificacion.modelo' if module == 'modelo' else module, name)


def cargar_modelo(ruta: str) -> Clasificador:
    """Abre un modelo guardado, también los generados con versiones anteriores de la herramienta."""
    with open(ruta, 'rb') as f:
        try:
            return _CargadorCompatible(f).load()
        except (pickle.UnpicklingError, EOFError, AttributeError, ModuleNotFoundError) as e:
            nombre = os.path.basename(ruta)
            raise ValueError(f'"{nombre}" no es un modelo válido. Vuelve a entrenarlo con historico.csv.') from e


def guardar_modelo(modelo: Clasificador, ruta: str) -> None:
    with open(ruta, 'wb') as f:
        pickle.dump(modelo, f)
