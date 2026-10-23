"""Genera index.html (la app que corre en el navegador) a partir de web/plantilla.html y el código Python.

Uso: python construir_web.py
Volver a correrlo cada vez que cambie algún .py. El index.html resultante sirve igual publicado en
GitHub Pages que abierto con doble clic desde el disco compartido. No incluye ningún dato.
"""
import json, os

AQUI = os.path.dirname(os.path.abspath(__file__))
PYODIDE = '0.27.7'
FUENTES = ['comun.py', 'leer_libro.py', 'modelo.py', 'salida.py', 'entrenar.py', 'clasificar_mes.py',
           'actualizar.py', 'app_web.py']

fuentes = {f: open(os.path.join(AQUI, f), encoding='utf-8').read() for f in FUENTES}
html = open(os.path.join(AQUI, 'web', 'plantilla.html'), encoding='utf-8').read()
# '</' escapado para que ningún código Python pueda cerrar el <script> por accidente
html = html.replace('/*__FUENTES__*/', json.dumps(fuentes, ensure_ascii=False).replace('</', r'<\/'))
html = html.replace('__PYODIDE__', PYODIDE)
open(os.path.join(AQUI, 'index.html'), 'w', encoding='utf-8').write(html)
print(f'index.html generado ({len(html) // 1024} KB, Pyodide {PYODIDE}).')
