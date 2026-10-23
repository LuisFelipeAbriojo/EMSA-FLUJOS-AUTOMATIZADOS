"""Lee el libro bancos de una empresa (una hoja por cuenta) y devuelve solo los movimientos del mes pedido."""
import datetime as dt, re
import openpyxl, pandas as pd
from comun import norm, banco, moneda, empresa_de

COLS = ['FECHA', 'CONCEPTO', 'DEBE', 'HABER', 'SALDO']


def _fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        s = norm(v)
        m = re.fullmatch(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})', s)
        if m:
            d, mo, a = map(int, m.groups())
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


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(',', ''))
        except ValueError:
            return None
    return None


def leer_empresa(ruta, anio, mes, empresa=None):
    """Devuelve (movimientos DataFrame, cuentas DataFrame con saldos)."""
    empresa = empresa or empresa_de(ruta.split('/')[-1])
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ini_mes = dt.date(anio, mes, 1)
    fin_mes = (dt.date(anio + (mes == 12), mes % 12 + 1, 1))
    movs, cuentas = [], []
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        cab = None
        for i, r in enumerate(filas[:15]):
            n = [norm(x) if x is not None else '' for x in r]
            if 'FECHA' in n and 'CONCEPTO' in n and 'DEBE' in n:
                cab = (i, n)
                break
        if not cab:
            continue  # hojas de saldos, cheques, tipo de cambio, etc.
        i0, n = cab
        c = {k: n.index(k) for k in COLS if k in n}
        cch = next((j for j, x in enumerate(n) if x.startswith('CH/')), None)
        cop = next((j for j, x in enumerate(n) if 'OPERACION' in x), None)
        texto_cab = ' '.join(norm(x) for r in filas[:i0] for x in r if x is not None)
        cuenta = next((m.group(0) for m in [re.search(r'(CTA|CUENTA)[^/]*?N\S*\s*[\d\- ]{6,}', texto_cab)] if m), '')
        mon = moneda(ws.title, texto_cab)
        saldo_ini, saldo_fin, n_mov, deb, hab = None, None, 0, 0.0, 0.0
        vacias, suelto = 0, False   # filas debajo de un bloque vacío = posibles proyectados
        for k, r in enumerate(filas[i0 + 1:], start=i0 + 2):
            r = list(r) + [None] * 20
            f = _fecha(r[c['FECHA']])
            if f is None:
                if not r[c['CONCEPTO']] and not _num(r[c['DEBE']]) and not _num(r[c['HABER']]):
                    vacias += 1
                continue
            if vacias >= 3 and saldo_ini is not None:
                suelto = True
            vacias = 0
            s = _num(r[c['SALDO']]) if 'SALDO' in c else None
            if f < ini_mes:
                if s is not None:
                    saldo_ini = s
                continue
            if f >= fin_mes:
                continue
            concepto = str(r[c['CONCEPTO']] or '').strip()
            d, h = _num(r[c['DEBE']]), _num(r[c['HABER']])
            if s is not None and not suelto:
                saldo_fin = s
            if norm(concepto).startswith('CIERRE AL'):
                continue
            base = dict(empresa=empresa, hoja=ws.title, cuenta=cuenta, banco=banco(ws.title), moneda=mon,
                        fecha=f, concepto=concepto, cheque=r[cch] if cch is not None else None,
                        operacion=r[cop] if cop is not None else None, fila_origen=k,
                        aviso='Fila suelta al final de la hoja (¿proyectado?)' if suelto else '')
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
                            saldo_final=saldo_fin or 0.0, movimientos=n_mov))
    return pd.DataFrame(movs), pd.DataFrame(cuentas)
