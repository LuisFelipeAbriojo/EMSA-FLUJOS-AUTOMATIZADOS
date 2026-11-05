"""Lee el libro bancos de una empresa (una hoja por cuenta) y devuelve solo los movimientos del mes pedido."""
from __future__ import annotations

import datetime as dt
import re
from pathlib import PurePath
from typing import Any

import openpyxl
import pandas as pd

from flujos.lectura.normalizar import banco, empresa_de, es_overnight, moneda, norm

COLS = ['FECHA', 'CONCEPTO', 'DEBE', 'HABER', 'SALDO']
# ELSA llama de otra forma a la columna del concepto
ALIAS_CONCEPTO = ('DESCRIPCION', 'CLIENTE')


def _columna_concepto(n: list[str]) -> int | None:
    if 'CONCEPTO' in n:
        return n.index('CONCEPTO')
    return next((j for j, x in enumerate(n) if x.startswith(ALIAS_CONCEPTO)), None)


def leer_fecha(v: Any) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        s = norm(v)
        m = re.fullmatch(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})', s)
        al_reves = re.fullmatch(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', s)  # 2026.09.01
        if m:
            d, mo, a = map(int, m.groups())
        elif al_reves:
            a, mo, d = map(int, al_reves.groups())
        else:  # fechas escritas a mano: 1-SET-26, 26-APR-29, 5 OCT 2026
            m = re.fullmatch(r'(\d{1,2})[\-/ .]*([A-Z]{3})[A-Z]*[\-/ .]*(\d{2,4})', s)
            if not m or m.group(2) not in _MESES:
                return None
            d, mo, a = int(m.group(1)), _MESES[m.group(2)], int(m.group(3))
        a += 2000 if a < 100 else 0
        try:
            return dt.date(a, mo, d)
        except ValueError:
            return None
    return None


_MESES = {'ENE': 1, 'JAN': 1, 'FEB': 2, 'MAR': 3, 'ABR': 4, 'APR': 4, 'MAY': 5, 'JUN': 6, 'JUL': 7,
          'AGO': 8, 'AUG': 8, 'SET': 9, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DIC': 12, 'DEC': 12}


def _num(v: Any) -> float | None:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(',', ''))
        except ValueError:
            return None
    return None


def leer_empresa(ruta: str, anio: int, mes: int, empresa: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (movimientos DataFrame, cuentas DataFrame con saldos)."""
    mov, cuentas, _ = leer_con_overnights(ruta, anio, mes, empresa)
    return mov, cuentas


def leer_con_overnights(ruta: str, anio: int, mes: int,
                        empresa: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Como leer_empresa, y además los overnights de este mes y sus vecinos: los cargos y abonos pueden cruzar de mes."""
    empresa = empresa or empresa_de(PurePath(ruta).name)
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ini_mes = dt.date(anio, mes, 1)
    fin_mes = (dt.date(anio + (mes == 12), mes % 12 + 1, 1))
    movs, cuentas, overnights = [], [], []
    desde = dt.date(anio - (mes == 1), (mes - 2) % 12 + 1, 1)
    hasta = dt.date(fin_mes.year + (fin_mes.month == 12), fin_mes.month % 12 + 1, 1)
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        cab = None
        for i, r in enumerate(filas[:15]):
            n = [norm(x) if x is not None else '' for x in r]
            if 'FECHA' in n and 'DEBE' in n and _columna_concepto(n) is not None:
                cab = (i, n)
                break
        if not cab:
            continue  # hojas de saldos, cheques, tipo de cambio, etc.
        i0, n = cab
        c = {k: n.index(k) for k in COLS if k in n}
        c['CONCEPTO'] = _columna_concepto(n)
        cch = next((j for j, x in enumerate(n) if x.startswith('CH/')), None)
        cop = next((j for j, x in enumerate(n) if 'OPERACION' in x), None)
        texto_cab = ' '.join(norm(x) for r in filas[:i0] for x in r if x is not None)
        cuenta = next((m.group(0) for m in [re.search(r'(CTA|CUENTA)[^/]*?N\S*\s*[\d\- ]{6,}', texto_cab)] if m), '')
        mon = moneda(ws.title, texto_cab)
        saldo_ini, saldo_fin, n_mov, deb, hab = None, None, 0, 0.0, 0.0
        visto_mes = False  # el saldo inicial es el de la fila justo antes del primer movimiento del mes
        vacias, suelto = 0, False   # filas debajo de un bloque vacío = posibles proyectados
        # 1ª pasada: filas con fecha, para detectar fechas fuera de lugar
        fechadas = []
        for k, r in enumerate(filas[i0 + 1:], start=i0 + 2):
            r = list(r) + [None] * 20
            f = leer_fecha(r[c['FECHA']])
            if f is None:
                if not r[c['CONCEPTO']] and not _num(r[c['DEBE']]) and not _num(r[c['HABER']]):
                    vacias += 1
                continue
            fechadas.append([k, r, f, vacias])
            vacias = 0
        # Una fila cuyo mes no coincide con el de las filas de arriba y de abajo (que sí coinciden entre sí)
        # está mal fechada (ej. 08/09 en vez de 09/08). El saldo del libro la cuenta donde está ubicada,
        # así que se asigna al mes de sus vecinas y se avisa.
        # También bloques: varias filas seguidas con otro mes (ej. 17 filas "21/08" en medio del 21/09).
        def ym(x: dt.date) -> int:
            return x.year * 12 + x.month

        alertas = []
        tramos = []  # [inicio, fin) de filas consecutivas con el mismo mes
        for i, it in enumerate(fechadas):
            if tramos and ym(fechadas[tramos[-1][0]][2]) == ym(it[2]):
                tramos[-1][1] = i + 1
            else:
                tramos.append([i, i + 1])
        for t in range(1, len(tramos) - 1):
            (pa, pb), (a0, a1), (na, nb) = tramos[t - 1], tramos[t], tramos[t + 1]
            antes, despues = fechadas[pb - 1][2], fechadas[na][2]
            largo = a1 - a0
            if ym(antes) == ym(despues) and largo <= 40 and largo < (pb - pa) + (nb - na):
                for i in range(a0, a1):
                    f = fechadas[i][2]
                    fechadas[i].append(f)              # fecha original
                    fechadas[i][2] = antes             # el mes que le da su ubicación en el libro
                f = fechadas[a0][4]
                if ini_mes <= f < fin_mes or ini_mes <= antes < fin_mes:
                    filas_txt = (f'fila {fechadas[a0][0]}' if largo == 1
                                 else f'filas {fechadas[a0][0]}-{fechadas[a1 - 1][0]} ({largo} movimientos)')
                    alertas.append(f'{ws.title} {filas_txt}: fecha {f:%d/%m/%Y} entre movimientos del '
                                   f'{antes:%d/%m/%Y}; se toma como {antes:%m/%Y} por su ubicación '
                                   f'(¿mes mal digitado?)')
        for item in fechadas:
            k, r, f, vac = item[:4]
            f_orig = item[4] if len(item) > 4 else None
            if desde <= f < hasta and es_overnight(r[c['CONCEPTO']] or ''):
                debe, haber = _num(r[c['DEBE']]), _num(r[c['HABER']])
                for tipo, monto in (('INGRESO', debe), ('EGRESO', haber)):
                    if monto and monto > 1:
                        overnights.append(dict(hoja=ws.title, banco=banco(ws.title), moneda=mon, fecha=f,
                                               concepto=str(r[c['CONCEPTO']]).strip(), tipo=tipo, monto=round(monto, 2)))
            if vac >= 3 and saldo_ini is not None:
                suelto = True
            s = _num(r[c['SALDO']]) if 'SALDO' in c else None
            if f < ini_mes:
                if s is not None and not visto_mes:
                    saldo_ini = s
                continue
            if f >= fin_mes:
                continue
            visto_mes = True
            concepto = str(r[c['CONCEPTO']] or '').strip()
            d, h = _num(r[c['DEBE']]), _num(r[c['HABER']])
            if s is not None and not suelto:
                saldo_fin = s
            if norm(concepto).startswith('CIERRE AL'):
                continue
            aviso = 'Fila suelta al final de la hoja (¿proyectado?)' if suelto else ''
            if f_orig:
                aviso = (aviso + f' Fecha en el libro: {f_orig:%d/%m/%Y}, fuera de orden; se toma como de este mes '
                                 f'por su ubicación (revisar la fecha).').strip()
            base = dict(empresa=empresa, hoja=ws.title, cuenta=cuenta, banco=banco(ws.title), moneda=mon,
                        fecha=f_orig or f, concepto=concepto, cheque=r[cch] if cch is not None else None,
                        operacion=r[cop] if cop is not None else None, fila_origen=k, aviso=aviso)
            # BBVA y otros: el ITF viene en la misma fila que el movimiento (DEBE y HABER a la vez, ~0.005%)
            itf = None
            if d and h and min(d, h) / max(d, h) <= 0.0001:
                itf = 'EGRESO' if h < d else 'INGRESO'
            for tipo, monto in (('INGRESO', d), ('EGRESO', h)):
                if monto:
                    extra = {'concepto': 'ITF', 'itf_misma_fila': 1, 'concepto_padre': concepto} if tipo == itf \
                        else {'itf_misma_fila': 0, 'concepto_padre': ''}
                    movs.append({**base, 'tipo': tipo, 'monto': round(monto, 2), **extra})
                    n_mov += 1
                    deb += monto if tipo == 'INGRESO' else 0
                    hab += monto if tipo == 'EGRESO' else 0
        if saldo_fin is None:
            saldo_fin = saldo_ini
        cuentas.append(dict(empresa=empresa, hoja=ws.title, cuenta=cuenta, banco=banco(ws.title), moneda=mon,
                            saldo_inicial=saldo_ini or 0.0, ingresos=round(deb, 2), egresos=round(hab, 2),
                            saldo_final=saldo_fin or 0.0, movimientos=n_mov, alertas=' | '.join(alertas)))
    return pd.DataFrame(movs), pd.DataFrame(cuentas), pd.DataFrame(overnights)
