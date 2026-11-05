"""Guarda el flujo con los valores calculados dentro del archivo.

openpyxl guarda las fórmulas sin su resultado: en la Vista protegida de Excel (la que abre todo archivo descargado)
se verían vacías hasta pulsar "Habilitar edición". Por eso se escriben los valores que calculó el evaluador.
Al habilitar la edición, Excel vuelve a calcular todo.
"""
from __future__ import annotations

import io
import re
import zipfile
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.utils import coordinate_to_tuple

from flujos.validacion.evaluador import ErrorExcel, Evaluador

_CELDA_FORMULA = re.compile(r'<c r="([A-Z]+\d+)"([^>]*)>(<f>.*?</f>|<f [^>]*>.*?</f>)<v\s*/>|<c r="([A-Z]+\d+)"([^>]*)>'
                            r'(<f>.*?</f>|<f [^>]*>.*?</f>)<v></v>', re.S)


def _archivos_de_hojas(z: zipfile.ZipFile) -> dict[str, str]:
    libro = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    destino = {}
    for m in re.finditer(r'<Relationship [^>]*>', rels):
        t = m.group(0)
        rid = re.search(r'Id="([^"]+)"', t).group(1)
        destino[rid] = re.search(r'Target="([^"]+)"', t).group(1).lstrip('/')
    hojas = {}
    for m in re.finditer(r'<sheet [^>]*>', libro):
        t = m.group(0)
        nombre = re.search(r'name="([^"]+)"', t).group(1)
        rid = re.search(r'r:id="([^"]+)"', t).group(1)
        ruta = destino[rid]
        hojas[ruta if ruta.startswith('xl/') else f'xl/{ruta}'] = _desescapar(nombre)
    return hojas


def _desescapar(s: str) -> str:
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')


def guardar_con_valores(wb: Workbook, ruta: str, ev: Evaluador | None = None) -> None:
    ev = ev or Evaluador(wb)
    temporal = io.BytesIO()
    wb.save(temporal)
    temporal.seek(0)
    with zipfile.ZipFile(temporal) as entrada, zipfile.ZipFile(ruta, 'w', zipfile.ZIP_DEFLATED) as salida:
        hojas = _archivos_de_hojas(entrada)
        for item in entrada.infolist():
            datos = entrada.read(item.filename)
            hoja = hojas.get(item.filename)
            if hoja is not None:
                datos = _con_valores(datos.decode('utf-8'), hoja, ev).encode('utf-8')
            salida.writestr(item, datos)


def _con_valores(xml: str, hoja: str, ev: Evaluador) -> str:
    def cambia(m: re.Match) -> str:
        coord, attrs, formula = (m.group(1), m.group(2), m.group(3)) if m.group(1) else (m.group(4), m.group(5), m.group(6))
        fila, col = coordinate_to_tuple(coord)
        try:
            v = ev._valor_en(hoja, fila, col)
        except (ValueError, ZeroDivisionError, OverflowError):
            return m.group(0)
        attrs = re.sub(r'\st="[^"]*"', '', attrs)
        if isinstance(v, ErrorExcel):
            return f'<c r="{coord}"{attrs} t="e">{formula}<v>{escape(v.codigo)}</v></c>'[:-4]
        if isinstance(v, bool):
            return f'<c r="{coord}"{attrs} t="b">{formula}<v>{int(v)}</v>'
        if isinstance(v, (int, float)):
            return f'<c r="{coord}"{attrs}>{formula}<v>{v!r}</v>'
        if v is None:
            return f'<c r="{coord}"{attrs}>{formula}<v>0</v>'
        return f'<c r="{coord}"{attrs} t="str">{formula}<v>{escape(str(v))}</v>'

    return _CELDA_FORMULA.sub(cambia, xml)
