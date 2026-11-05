"""Genera dist/index.html: la app completa en un solo archivo, sin ningún dato.

Uso: python scripts/construir_web.py

Incrusta la interfaz, el paquete `flujos` como wheel y openpyxl, de modo que la misma página sirve publicada
en GitHub Pages o abierta con doble clic desde el disco compartido, sin depender de pypi.org.
"""
from __future__ import annotations

import base64
import glob
import json
import os
import shutil
import subprocess
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PYODIDE = '0.27.7'
TERCEROS = ['openpyxl==3.1.5', 'et_xmlfile==2.0.0']


def _wheels() -> list[str]:
    carpeta = os.path.join(RAIZ, 'build', 'web')
    shutil.rmtree(carpeta, ignore_errors=True)
    subprocess.run([sys.executable, '-m', 'build', '--wheel', '--outdir', carpeta, RAIZ], check=True,
                   stdout=subprocess.DEVNULL)
    subprocess.run([sys.executable, '-m', 'pip', 'download', '-q', '--no-deps', '--only-binary', ':all:',
                    '-d', carpeta, *TERCEROS], check=True)
    rutas = sorted(glob.glob(os.path.join(carpeta, '*.whl')))
    if len(rutas) != 1 + len(TERCEROS):
        raise SystemExit(f'Se esperaban {1 + len(TERCEROS)} wheels y hay {len(rutas)}.')
    return rutas


def _leer(nombre: str) -> str:
    with open(os.path.join(RAIZ, 'web', nombre), encoding='utf-8') as f:
        return f.read()


def main() -> int:
    wheels = []
    for ruta in _wheels():
        with open(ruta, 'rb') as f:
            wheels.append(base64.b64encode(f.read()).decode())
    html = _leer('index.html')
    html = html.replace('/*__ESTILOS__*/', _leer('estilos.css'))
    html = html.replace('/*__WORKER__*/', _leer('worker.js').replace('__PYODIDE__', PYODIDE))
    html = html.replace('/*__APP__*/', _leer('app.js').replace('/*__WHEELS__*/', json.dumps(wheels)))
    os.makedirs(os.path.join(RAIZ, 'dist'), exist_ok=True)
    destino = os.path.join(RAIZ, 'dist', 'index.html')
    with open(destino, 'w', encoding='utf-8', newline='\n') as f:
        f.write(html)
    print(f'dist/index.html generado ({len(html) // 1024} KB, Pyodide {PYODIDE}).')
    return 0


if __name__ == '__main__':
    sys.exit(main())
