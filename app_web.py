"""Puente entre la página web (index.html) y el clasificador. Corre dentro del navegador con Pyodide.

Cada función recibe rutas del sistema de archivos virtual del navegador y deja lo que hay que
descargar en /salida. Nada se envía a ningún servidor.
"""
import os, pickle, shutil, warnings, zipfile
warnings.filterwarnings('ignore')
import pandas as pd
from entrenar import entrenar
from clasificar_mes import clasificar as _clasificar
from actualizar import actualizar as _actualizar

SALIDA = '/salida'
MODELO = None
HIST = None


def _vaciar_salida():
    shutil.rmtree(SALIDA, ignore_errors=True)
    os.makedirs(SALIDA)


def _info_modelo():
    m = MODELO
    return dict(movimientos=int(m.n_entrenamiento),
                empresas=sorted({e for e, _ in m.permitidas}),
                categorias=int(len(m.lr.classes_)),
                meses=sorted(HIST.periodo.astype(str).unique().tolist()) if HIST is not None else [])


def cargar_historico(ruta, avisar):
    """Entrena con historico.csv y deja modelo_web.pkl para descargar."""
    global MODELO, HIST
    _vaciar_salida()
    HIST = pd.read_csv(ruta)
    avisar(f'Histórico leído: {len(HIST):,} movimientos. Entrenando (puede tardar unos minutos)...')
    MODELO, _ = entrenar(HIST)
    pickle.dump(MODELO, open(f'{SALIDA}/modelo_web.pkl', 'wb'))
    avisar('Modelo entrenado.')
    return _info_modelo()


def cargar_modelo(ruta, avisar):
    global MODELO
    _vaciar_salida()
    MODELO = pickle.load(open(ruta, 'rb'))
    avisar('Modelo cargado.')
    return _info_modelo()


def clasificar(rutas, anio, mes, umbral, avisar):
    _vaciar_salida()
    rutas = [r for r in rutas if not os.path.basename(r).startswith('~$')]
    hechos = _clasificar(rutas, int(anio), int(mes), MODELO, SALIDA, float(umbral), avisar)
    for h in hechos:
        h['archivo'] = os.path.basename(h.pop('ruta'))
    if len(hechos) > 1:
        with zipfile.ZipFile(f'{SALIDA}/FLUJOS CLASIFICADOS {anio}-{int(mes):02d}.zip', 'w', zipfile.ZIP_DEFLATED) as z:
            for h in hechos:
                z.write(f"{SALIDA}/{h['archivo']}", h['archivo'])
    return hechos


def actualizar(rutas, ruta_historico, umbral, avisar):
    """Rehace los Excel revisados y devuelve el histórico con esos meses agregados."""
    global HIST
    _vaciar_salida()
    if ruta_historico:
        HIST = pd.read_csv(ruta_historico)
    if HIST is None:
        raise ValueError('Falta historico.csv: cárgalo en el paso 1 o junto con los Excel revisados.')
    copias = []
    for r in rutas:
        destino = f'{SALIDA}/{os.path.basename(r)}'
        shutil.copy(r, destino)
        copias.append(destino)
    HIST = _actualizar(copias, HIST, MODELO, float(umbral), avisar)
    HIST.to_csv(f'{SALIDA}/historico.csv', index=False)
    avisar(f'Histórico actualizado: {len(HIST):,} movimientos.')
    return dict(movimientos=len(HIST))


def reentrenar(avisar):
    global MODELO
    _vaciar_salida()
    avisar('Reentrenando con el histórico actualizado...')
    MODELO, _ = entrenar(HIST)
    pickle.dump(MODELO, open(f'{SALIDA}/modelo_web.pkl', 'wb'))
    avisar('Modelo reentrenado.')
    return _info_modelo()
