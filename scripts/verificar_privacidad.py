"""Bloquea datos de las empresas antes de que lleguen al repositorio público.

Uso:
    python scripts/verificar_privacidad.py          # lo que está en el área de preparación (pre-commit)
    python scripts/verificar_privacidad.py --todo   # todos los archivos versionados (CI)
"""
from __future__ import annotations

import re
import subprocess
import sys

EXTENSIONES_PROHIBIDAS = {'.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.csv', '.tsv', '.pkl', '.pickle', '.joblib',
                          '.zip', '.7z', '.rar', '.parquet', '.pdf'}
TAMANO_MAXIMO = 500 * 1024
PATRONES = {
    'número de cuenta': re.compile(rb'(?<![\d.-])\d{3}-\d{5,}-\d(?:-\d{2})?(?![\d-])'),
    'RUC': re.compile(rb'(?<!\d)(?:10|20)\d{9}(?!\d)'),
    'CCI': re.compile(rb'(?<!\d)\d{20}(?!\d)'),
}
# index.html generado incrusta wheels en base64 y nunca se revisa como texto
SIN_REVISAR_CONTENIDO = {'index.html'}


def _git(*args: str) -> bytes:
    return subprocess.run(['git', *args], capture_output=True, check=True).stdout


def _rutas(todo: bool) -> list[str]:
    salida = _git('ls-files', '-z') if todo else _git('diff', '--cached', '--name-only', '--diff-filter=ACMR', '-z')
    return [r.decode() for r in salida.split(b'\0') if r]


def revisar(ruta: str, contenido: bytes) -> list[str]:
    """Devuelve los problemas encontrados en un archivo; lista vacía si está limpio."""
    problemas = []
    nombre = ruta.rsplit('/', 1)[-1].lower()
    if any(nombre.endswith(ext) for ext in EXTENSIONES_PROHIBIDAS):
        problemas.append('tipo de archivo de datos')
    if len(contenido) > TAMANO_MAXIMO:
        problemas.append(f'archivo grande ({len(contenido) // 1024} KB)')
    if ruta not in SIN_REVISAR_CONTENIDO and b'\0' not in contenido:
        for que, patron in PATRONES.items():
            if patron.search(contenido):
                problemas.append(f'parece contener un {que}')
    return problemas


def main() -> int:
    todo = '--todo' in sys.argv
    errores = 0
    for ruta in _rutas(todo):
        contenido = open(ruta, 'rb').read() if todo else _git('show', f':{ruta}')
        for problema in revisar(ruta, contenido):
            print(f'BLOQUEADO  {ruta}: {problema}')
            errores += 1
    if errores:
        print('\nEste repositorio es público. Saca esos datos antes de continuar.')
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
