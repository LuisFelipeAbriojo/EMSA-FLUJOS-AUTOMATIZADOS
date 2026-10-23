"""Después de revisar: rehace las hojas por cuenta y guarda las correcciones para que el modelo aprenda.

Uso: python actualizar.py "ICP FLUJO SETIEMBRE 2026 CLASIFICADO.xlsx" [...] [--historico historico.csv]

Lee la hoja MOVIMIENTOS (con las categorías que corrigieron), regenera las hojas agrupadas por cuenta
y agrega el mes al histórico. Luego: python entrenar.py historico.csv
"""
import argparse, os, pickle, re, sys, warnings
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings('ignore')
import openpyxl, pandas as pd
from comun import banco, moneda
from salida import escribir, MESES

AQUI = os.path.dirname(os.path.abspath(__file__))


def actualizar(archivos, h, modelo, umbral=0.9, avisar=print):
    """Rehace cada archivo revisado en su sitio y devuelve el histórico h con esos meses agregados."""
    nuevos = []
    for f in archivos:
        emp = os.path.basename(f).split()[0]
        m = re.search(r'FLUJO (\w+) (\d{4})', os.path.basename(f))
        mes, anio = MESES.index(m.group(1)), int(m.group(2))
        mv = pd.read_excel(f, sheet_name='MOVIMIENTOS')
        # saldos y cuentas desde la hoja RESUMEN
        wb = openpyxl.load_workbook(f, data_only=False)
        ws = wb['RESUMEN']
        cuentas = []
        for r in ws.iter_rows(min_row=5, values_only=True):
            if not r[0] or r[2] not in ('USD', 'PEN'):
                break
            cuentas.append(dict(hoja=r[0], cuenta=r[1], moneda=r[2], saldo_inicial=r[3], ingresos=r[4], egresos=r[5],
                                saldo_final=r[7], movimientos=r[9]))
        cuentas = pd.DataFrame(cuentas)
        mov = pd.DataFrame(dict(
            empresa=emp, hoja=mv.CUENTA, moneda=mv.MONEDA, fecha=pd.to_datetime(mv.FECHA).dt.date, concepto=mv.CONCEPTO,
            tipo=mv.TIPO, monto=mv.MONTO, categoria=mv.CATEGORIA.astype(str).str.strip().str.upper(),
            confianza=mv.CONFIANZA, alternativa=mv.ALTERNATIVA, origen=mv.ORIGEN, aviso=mv.NOTA.fillna(''),
            operacion=mv['Nº OPERACION'], cheque=mv['CH/Nº'], fila_origen=mv['FILA EN LIBRO'],
            itf_misma_fila=mv.NOTA.fillna('').str.contains('misma fila').astype(int), concepto_padre=''))
        # lo que el equipo revisó ya no está "por revisar"
        mov.loc[mov.confianza < umbral, ['confianza', 'origen']] = [1.0, 'revisado']
        escribir(f, emp, anio, mes, mov, cuentas, modelo, umbral)
        nuevos.append(pd.DataFrame(dict(
            itf_misma_fila=mov.itf_misma_fila, empresa=emp, periodo=f'{anio}-{mes:02d}', hoja=mov.hoja,
            banco=mov.hoja.map(banco), moneda=mov.moneda, fecha=mov.fecha.astype(str), concepto=mov.concepto,
            tipo=mov.tipo, monto=mov.monto, categoria=mov.categoria,
            orden=[modelo.orden_de(emp, t, c) for t, c in zip(mov.tipo, mov.categoria)],
            archivo=os.path.basename(f), fuente='revision')))
        avisar(f'{emp}: hojas por cuenta rehechas; {len(mov)} movimientos listos para el histórico.')

    h = h.copy()
    if 'fuente' not in h:
        h['fuente'] = 'flujo cerrado'
    n = pd.concat(nuevos)
    clave = set(zip(n.empresa, n.periodo))
    h = h[~pd.Series(list(zip(h.empresa, h.periodo)), dtype=object).isin(clave).values | (h.fuente != 'revision')]
    return pd.concat([h, n])


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('archivos', nargs='+')
    ap.add_argument('--historico', default=os.path.join(AQUI, 'historico.csv'))
    ap.add_argument('--modelo', default=os.path.join(AQUI, 'modelo.pkl'))
    ap.add_argument('--umbral', type=float, default=0.9)
    a = ap.parse_args()
    modelo = pickle.load(open(a.modelo, 'rb'))
    h = pd.read_csv(a.historico) if os.path.exists(a.historico) else pd.DataFrame()
    actualizar(a.archivos, h, modelo, a.umbral).to_csv(a.historico, index=False)
    print(f'Histórico actualizado: {a.historico}. Ahora corre: python entrenar.py "{a.historico}"')
