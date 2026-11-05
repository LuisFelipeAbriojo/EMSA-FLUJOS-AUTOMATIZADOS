"""De dónde salen los movimientos del armado: el Excel clasificado y revisado de la empresa."""
from __future__ import annotations

import datetime as dt
import logging
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

import openpyxl
import pandas as pd

from flujos.armado.estructura import Bloque, Estructura, Fila, compensa_intereses, interes_va_en_abono
from flujos.armado.overnight import Overnight, emparejar
from flujos.clasificacion.reglas import secciones_por_operacion, tipo_operacion
from flujos.lectura.normalizar import banco, categoria, es_overnight, moneda, norm

log = logging.getLogger(__name__)
HOJA_OVERNIGHTS = 'OVERNIGHTS'
PREFIJO_INTERESES = 'INTERESES OVERNIGHT'


@dataclass
class Clasificado:
    movimientos: pd.DataFrame
    cuentas: pd.DataFrame
    overnights: pd.DataFrame
    alertas: list[str] = field(default_factory=list)


def leer_clasificado(ruta: str) -> Clasificado:
    """Lee las categorías ya revisadas, los saldos por cuenta, los overnights y las fechas sospechosas."""
    wb = openpyxl.load_workbook(ruta, data_only=False)
    faltan = [h for h in ('MOVIMIENTOS', 'RESUMEN') if h not in wb.sheetnames]
    if faltan:
        raise ValueError(f'"{ruta.rsplit("/", 1)[-1]}" no es un Excel clasificado: le falta la hoja {faltan[0]}.')
    mv = pd.read_excel(ruta, sheet_name='MOVIMIENTOS')
    mov = pd.DataFrame(dict(
        hoja=mv.CUENTA, moneda=mv.MONEDA, fecha=pd.to_datetime(mv.FECHA).dt.date, concepto=mv.CONCEPTO, tipo=mv.TIPO,
        monto=mv.MONTO, categoria=mv.CATEGORIA.astype(str).map(categoria), revisar=mv.REVISAR.fillna('') == 'SI',
        nota=mv.NOTA.fillna(''), fila=mv['FILA EN LIBRO']))
    filas = []
    alertas = []
    for r in wb['RESUMEN'].iter_rows(min_row=5, values_only=True):
        if r[0] and r[2] in ('USD', 'PEN'):
            filas.append(dict(hoja=r[0], cuenta=r[1], moneda=r[2], saldo_inicial=r[3], ingresos=r[4], egresos=r[5],
                              saldo_final=r[7], movimientos=r[9]))
        elif isinstance(r[0], str) and r[0].startswith('• '):
            alertas.append(r[0][2:])
    overnights = pd.DataFrame()
    if HOJA_OVERNIGHTS in wb.sheetnames:
        overnights = pd.read_excel(ruta, sheet_name=HOJA_OVERNIGHTS)
        if not overnights.empty:
            overnights['fecha'] = pd.to_datetime(overnights.fecha).dt.date
    return Clasificado(mov, pd.DataFrame(filas), overnights, alertas)


def overnights_del_clasificado(c: Clasificado) -> list[Overnight]:
    return emparejar(c.overnights)


def asignar_hojas(cuentas: pd.DataFrame, hojas_flujo: dict[str, float | None]) -> tuple[dict[str, list[str]], list[str]]:
    """Empareja cada cuenta del libro con su hoja del flujo.

    Los nombres no coinciden (BCP S / BCP SOLES, SBK / SCOT). Se usa banco y moneda; si hay varias candidatas, manda
    la continuidad de saldos: el saldo inicial del mes es el saldo final de la hoja en el mes anterior.
    Devuelve (hoja del flujo -> cuentas del libro, avisos).
    """
    candidatas: dict[tuple[str, str], list[str]] = defaultdict(list)
    for h in hojas_flujo:
        candidatas[(banco(h), moneda(h))].append(h)
    asignadas: dict[str, list[str]] = defaultdict(list)
    avisos = []
    fm = [h for h in hojas_flujo if banco(h) == 'FM']
    for _, c in cuentas.iterrows():
        if banco(c.hoja) == 'FM' and fm:  # los fondos del libro van a una sola hoja del flujo
            asignadas[fm[0]].append(c.hoja)
            continue
        opciones = candidatas.get((banco(c.hoja), c.moneda), [])
        if len(opciones) > 1:
            por_saldo = [h for h in opciones if hojas_flujo[h] is not None
                         and abs(hojas_flujo[h] - (c.saldo_inicial or 0)) < 0.05]
            opciones = por_saldo if len(por_saldo) == 1 else opciones
        if len(opciones) == 1:
            asignadas[opciones[0]].append(c.hoja)
        elif c.movimientos:
            avisos.append(f'La cuenta "{c.hoja}" del libro no tiene una hoja clara en el flujo '
                          f'({len(opciones)} candidatas). Sus movimientos no se incluyeron.')
    return dict(asignadas), avisos


def _orden(fila: dict[str, Any], por_operacion: bool) -> tuple:
    return (tipo_operacion(fila['concepto']) if por_operacion else 0, fila['fecha'], fila['fila'])


def _como_datetime(f: Any) -> Any:
    return dt.datetime.combine(f, dt.time()) if isinstance(f, dt.date) and not isinstance(f, dt.datetime) else f


def _categoria_fm(r: dict[str, Any], e: Estructura) -> str:
    """En Fondos Mutuos la sección sale del lado y del concepto; cada hoja la llama a su manera."""
    if r['tipo'] == 'EGRESO':
        opciones = ['CARGO TRANSF. CTAS', 'CREDICORP']
    elif re.search(r'INTERES', norm(r['concepto'])):
        opciones = ['OTROS INGRESOS', 'INTERFONDOS', 'CREDICORP']
    else:
        opciones = ['ABONO TRANSF. CTAS', 'CREDICORP']
    categorias = {b.categoria for b in e.bloques if b.tipo == r['tipo']}
    return next((o for o in opciones if o in categorias), r['categoria'])


def repartir(mov: pd.DataFrame, e: Estructura, vinculadas: set[int], overs: list[Overnight],
             hojas_libro: list[str], ws: Any = None,
             periodo: tuple[int, int] | None = None) -> tuple[list[list[Fila]], pd.DataFrame, list[str]]:
    """Reparte los movimientos en los bloques de la hoja. Devuelve (contenido, los que no tienen sección, avisos)."""
    por_operacion = secciones_por_operacion()
    avisos = []
    contenido: list[list[Fila]] = [[] for _ in e.bloques]

    def elegir(cat: str, tipo: str) -> int | None:
        iguales = [i for i, b in enumerate(e.bloques) if b.categoria == cat and b.tipo == tipo]
        if not iguales:  # una devolución de ITF va a la sección ITF aunque entre por el otro lado
            iguales = [i for i, b in enumerate(e.bloques) if b.categoria == cat]
        if not iguales:
            return None
        # si la sección está repetida, va a la que alimenta la hoja mensual
        return next((i for i in iguales if e.bloques[i].ultima in vinculadas), iguales[0])

    sin_seccion = []
    registros = mov.to_dict('records')
    es_fm = banco(e.hoja) == 'FM'
    abonos: dict[tuple[str, Any, float], tuple[Fila, int]] = {}
    for r in sorted(registros, key=lambda x: _orden(x, x['categoria'] in por_operacion)):
        cat = _categoria_fm(r, e) if es_fm else r['categoria']
        i = elegir(cat, r['tipo'])
        if i is None:
            sin_seccion.append(r)
            continue
        fila = Fila(_como_datetime(r['fecha']), r['concepto'], r['monto'] if r['tipo'] == 'INGRESO' else None,
                    r['monto'] if r['tipo'] == 'EGRESO' else None)
        if es_fm and e.col_fecha > 1:
            fila.otros[1] = re.sub(r'^FM\s*', '', str(r['hoja'])).upper()
        contenido[i].append(fila)
        if r['tipo'] == 'INGRESO':
            abonos[(r['hoja'], r['fecha'], round(r['monto'], 2))] = (fila, i)
    # interés = abono - cargo, en su propia sección
    intereses = _bloque_intereses(e.bloques, vinculadas)
    def del_mes(f: Any) -> bool:
        # las filas mal fechadas conservan su fecha original: el mes se compara con el periodo, no con esas fechas
        return f is not None and (periodo is None or (f.year, f.month) == periodo)

    con_cargo = {(o.hoja, o.fecha_abono, round(o.abono, 2)) for o in overs if o.abono is not None}
    for clave in abonos:
        fila = abonos[clave][0]
        if es_overnight(fila.concepto) and clave not in con_cargo:
            avisos.append(f'{e.hoja}: el abono "{fila.concepto}" del {clave[1]} no tiene su cargo en el libro: '
                          f'sin interés calculado. Revisar.')
    for o in overs:
        if o.hoja in hojas_libro and del_mes(o.fecha_abono) and o.interes:
            if intereses is None:
                avisos.append(f'{e.hoja}: no hay sección de intereses overnight para {o.concepto}.')
                continue
            par = abonos.get((o.hoja, o.fecha_abono, round(o.abono, 2)))
            # cada hoja registra el interés a su manera; la plantilla dice cuál
            if ws is not None and interes_va_en_abono(ws, e, e.bloques[intereses]):
                # la sección de intereses solo trae un total: el interés va en HABER, en la misma fila del abono
                if par is not None:
                    par[0].haber = o.interes
                else:
                    avisos.append(f'{e.hoja}: no encontré el abono de {o.concepto} para registrar su interés.')
                continue
            contenido[intereses].append(Fila(_como_datetime(o.fecha_abono), o.concepto, o.interes, None))
            # si tampoco se compensa con una celda, el abono queda solo como capital
            if par is not None and ws is not None and not compensa_intereses(ws, e, e.bloques[par[1]]):
                par[0].debe = round(par[0].debe - o.interes, 2)
    if intereses is not None:
        contenido[intereses].sort(key=lambda f: f.fecha)
    return contenido, pd.DataFrame(sin_seccion), avisos


def _bloque_intereses(bloques: list[Bloque], vinculadas: set[int]) -> int | None:
    iguales = [i for i, b in enumerate(bloques) if b.categoria.startswith(PREFIJO_INTERESES)]
    if not iguales:
        return None
    return next((i for i in iguales if bloques[i].ultima in vinculadas), iguales[0])


def nombre_salida(plantilla: str, mes_anterior: str, mes_nuevo: str) -> str:
    """Conserva el patrón de nombre de cada empresa: cambia solo el mes."""
    nombre = plantilla.rsplit('/', 1)[-1]
    base = norm(nombre)
    if mes_anterior in base:
        i = base.index(mes_anterior)
        return nombre[:i] + mes_nuevo + nombre[i + len(mes_anterior):]
    return f'FLUJO {mes_nuevo} {nombre}'
