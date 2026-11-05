"""El paso completo: Excel clasificado + flujo anterior + TC -> flujo del mes, notas y validaciones."""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Any

import openpyxl
import pandas as pd

from flujos.armado.armar import armar, hoja_mensual
from flujos.armado.estructura import Estructura, Fila, analizar
from flujos.armado.fuentes import asignar_hojas, leer_clasificado, nombre_salida, overnights_del_clasificado, repartir
from flujos.armado.guardar import guardar_con_valores
from flujos.armado.overnight import cruzan, hoja_over
from flujos.armado.rodar import ABREV, NOMBRES
from flujos.lectura.normalizar import norm
from flujos.reportes.notas import DatosNotas, escribir_notas, variaciones_total
from flujos.validacion.evaluador import Evaluador
from flujos.validacion.revision import (
    Control,
    errores_de_formula,
    errores_existentes,
    resumen_cuadra,
    saldos_contra_libro,
    secciones_contra_movimientos,
    subtotales_completos,
)

log = logging.getLogger(__name__)


@dataclass
class Armado:
    empresa: str
    flujo: str
    notas: str
    errores: int
    avisos: int
    movimientos: int
    sin_seccion: int


def _saldos_finales(plantilla: str) -> dict[str, float | None]:
    """Saldo final de cada hoja en el flujo anterior, para emparejar cuentas por continuidad."""
    wf = openpyxl.load_workbook(plantilla)
    wv = openpyxl.load_workbook(plantilla, data_only=True)
    mensual = hoja_mensual(wf)
    salida: dict[str, float | None] = {}
    for hoja in wf.sheetnames:
        if hoja in ('TOTAL', mensual, 'Resumen') or hoja == hoja_over(wf.sheetnames):
            continue
        e = analizar(wf[hoja])
        if e is None:
            continue
        salida[hoja] = None
        for fila in wf[hoja].iter_rows(max_row=e.fila_cabecera):
            for c in fila:
                if isinstance(c.value, str) and norm(c.value).startswith('SALDO FINAL'):
                    for dx in (-1, 1, 2):
                        if c.column + dx < 1:
                            continue
                        v = wv[hoja].cell(c.row, c.column + dx).value
                        if isinstance(v, (int, float)):
                            salida[hoja] = float(v)
                            break
    return salida


def _diferencia_de_cambio(wb: Any, ev: Evaluador, mensual: str) -> list[tuple[str, float]]:
    """Por cuenta en soles: saldo arrastrado por la hoja mensual menos saldo convertido al TC del mes."""
    ws = wb[mensual]
    fila_fin = None
    columna_rotulo = 1
    for col in (1, 2):
        for f in range(40, 95):
            v = ws.cell(f, col).value
            if isinstance(v, str) and norm(v).startswith('SALDO FINAL'):
                fila_fin, columna_rotulo = f, col
                break
        if fila_fin:
            break
    if fila_fin is None:
        return []
    cabecera = None
    for f in range(1, 12):
        if any(isinstance(ws.cell(f, c).value, str) and norm(ws.cell(f, c).value) in ('TOTALES', 'TOTAL')
               for c in range(1, 30)):
            cabecera = f
            break
    if cabecera is None:
        return []
    salida = []
    for c in range(columna_rotulo + 1, 30):
        nombre = ws.cell(cabecera, c).value
        if not isinstance(nombre, str) or norm(nombre) in ('TOTALES', 'TOTAL'):
            break
        a, b = ev._valor_en(mensual, fila_fin, c), ev._valor_en(mensual, fila_fin + 1, c)
        if ('S/' in nombre.upper() or 'SOL' in nombre.upper() or 'NAC' in nombre.upper()) and isinstance(a, (int, float)) \
                and isinstance(b, (int, float)) and abs(a - b) > 0.005:
            salida.append((' '.join(nombre.split()), round(a - b, 2)))
    return salida


def armar_flujo(clasificado: str, plantilla: str, empresa: str, anio: int, mes: int, tc: float, salida: str) -> Armado:
    """Arma el flujo y sus notas en la carpeta `salida`. No modifica los archivos de entrada."""
    if not 2 < tc < 6:
        raise ValueError(f'El tipo de cambio {tc} no parece válido: debe estar entre 2 y 6.')
    c = leer_clasificado(clasificado)
    saldos = _saldos_finales(plantilla)
    asignadas, avisos_cuentas = asignar_hojas(c.cuentas, saldos)
    overs = overnights_del_clasificado(c)
    salen, vuelven = cruzan(overs, anio, mes)
    sin_seccion: list[pd.DataFrame] = []
    avisos: list[str] = list(avisos_cuentas)

    def fuente(hoja: str, e: Estructura, vinculadas: set[int]) -> list[list[Fila]] | None:
        libros = asignadas.get(hoja)
        if not libros:
            return [[] for _ in e.bloques]
        contenido, resto, av = repartir(c.movimientos[c.movimientos.hoja.isin(libros)], e, vinculadas, overs, libros,
                                        plantilla_ws(hoja))
        avisos.extend(av)
        if not resto.empty:
            sin_seccion.append(resto)
        return contenido

    libro_plantilla = openpyxl.load_workbook(plantilla)
    previos = errores_existentes(libro_plantilla)

    def plantilla_ws(hoja: str) -> Any:
        return libro_plantilla[hoja]

    saldo_libro = {h: float(c.cuentas[c.cuentas.hoja.isin(ls)].saldo_final.fillna(0).sum()) for h, ls in asignadas.items()}
    res = armar(plantilla, empresa, anio, mes, tc, fuente, (salen, vuelven), saldo_libro)
    wb = res.libro
    r = res.rodaje
    # los errores previos ya se reportan con las hojas y filas del mes anterior
    ev = Evaluador(wb)
    faltantes = pd.concat(sin_seccion) if sin_seccion else pd.DataFrame()

    os.makedirs(salida, exist_ok=True)
    mes_anterior = 12 if mes == 1 else mes - 1
    plantilla_nombre = os.path.basename(plantilla)
    base = norm(plantilla_nombre)
    viejo = next((x for x in (NOMBRES[mes_anterior], 'SEPTIEMBRE' if mes_anterior == 9 else '', ABREV[mes_anterior])
                  if x and x in base), NOMBRES[mes_anterior])
    nombre = nombre_salida(plantilla_nombre, viejo, NOMBRES[mes])
    ruta_flujo = os.path.join(salida, nombre)
    guardar_con_valores(wb, ruta_flujo, ev)

    controles: list[Control] = []
    controles += saldos_contra_libro(wb, ev, asignadas, c.cuentas, saldos)
    intereses: dict[str, float] = {}
    for o in overs:
        if o.fecha_abono is not None and o.fecha_abono.year == anio and o.fecha_abono.month == mes and o.interes:
            intereses[o.hoja] = intereses.get(o.hoja, 0.0) + o.interes
    controles += secciones_contra_movimientos(wb, asignadas, c.movimientos, intereses)
    controles += subtotales_completos(wb, res.estructuras)
    controles += errores_de_formula(wb, ev, {(r.hoja_mensual_nueva if h == r.hoja_mensual_anterior else h, x) for h, x in previos})
    controles += resumen_cuadra(wb, ev, r.hoja_mensual_nueva)
    if not faltantes.empty:
        controles.append(Control('Todos los movimientos tienen sección', False,
                                 f'{len(faltantes)} movimientos sin sección: no están en el flujo. Ver Alertas.'))
    controles.append(Control('Sin filas por revisar', not c.movimientos.revisar.any(),
                             f'{int(c.movimientos.revisar.sum())} filas siguieron por revisar.', 'AVISO'))
    if r.tc_anterior is None:
        avisos.append('No encontré la celda del TC anterior en la hoja mensual.')

    tc_anterior = wb[r.hoja_mensual_nueva][r.tc_anterior].value if r.tc_anterior else None
    cuentas_sin = c.cuentas[c.cuentas.movimientos == 0].hoja.tolist() if not c.cuentas.empty else []
    datos = DatosNotas(
        empresa=empresa, mes_nombre=NOMBRES[mes], anio=anio, tc=tc, tc_anterior=tc_anterior, controles=controles,
        por_revisar=c.movimientos[c.movimientos.revisar], alertas_fechas=c.alertas, sin_seccion=faltantes,
        avisos=avisos + res.avisos + r.avisos, ajustes_quitados=r.ajustes_quitados,
        diferencia_cambio=_diferencia_de_cambio(wb, ev, r.hoja_mensual_nueva),
        variaciones=variaciones_total(wb['TOTAL'], ev, r.columna_nueva),
        cuentas_sin_movimientos=cuentas_sin, vinculos_nuevos=res.vinculos_nuevos)
    ruta_notas = os.path.join(salida, f'{empresa} NOTAS {NOMBRES[mes]} {anio}.xlsx')
    conteo = escribir_notas(ruta_notas, datos)
    log.info('%s: %s | errores %d, avisos %d, sin sección %d', empresa, nombre, conteo['errores'], conteo['avisos'], len(faltantes))
    return Armado(empresa, ruta_flujo, ruta_notas, conteo['errores'], conteo['avisos'], len(c.movimientos), len(faltantes))
