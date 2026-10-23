"""Clasifica el mes cerrado a partir de los libros bancos de cada empresa.

Uso:
    python clasificar_mes.py LIBROS... [--mes 2026-10] [--modelo modelo.pkl] [--salida carpeta] [--umbral 0.9]

LIBROS: los .xlsx de cada empresa (ECSA SET 2026.xlsx, ...) o una carpeta que los contenga.
--mes: por defecto el mes anterior al de hoy (en noviembre procesa octubre).
Genera "<EMPRESA> FLUJO <MES> CLASIFICADO.xlsx" por empresa.
"""
import argparse, datetime as dt, glob, os, pickle, sys, warnings
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings('ignore')
import pandas as pd
from leer_libro import leer_empresa
from comun import empresa_de
from salida import escribir, MESES

AQUI = os.path.dirname(os.path.abspath(__file__))


def mes_anterior():
    h = dt.date.today().replace(day=1) - dt.timedelta(days=1)
    return h.year, h.month


def clasificar(archivos, anio, mes, modelo, salida='.', umbral=0.9, avisar=print):
    """Escribe un Excel clasificado por empresa en `salida` y devuelve un resumen por archivo generado."""
    os.makedirs(salida, exist_ok=True)
    avisar(f'Mes a procesar: {MESES[mes]} {anio}  (umbral de revisión {umbral:.0%})\n')
    hechos = []
    for f in archivos:
        emp = empresa_de(os.path.basename(f))
        if not emp:
            avisar(f'  ?  No reconozco la empresa en "{os.path.basename(f)}", lo salto.')
            continue
        mov, cuentas = leer_empresa(f, anio, mes, emp)
        if mov.empty:
            avisar(f'  {emp:5s} sin movimientos en {MESES[mes]} {anio}')
            continue
        pred = modelo.predict(mov)
        mov = pd.concat([mov.reset_index(drop=True), pred], axis=1)
        ruta = os.path.join(salida, f'{emp} FLUJO {MESES[mes]} {anio} CLASIFICADO.xlsx')
        mov = escribir(ruta, emp, anio, mes, mov, cuentas, modelo, umbral)
        descuadre = cuentas[(cuentas.saldo_inicial + cuentas.ingresos - cuentas.egresos - cuentas.saldo_final).abs() > 0.05]
        hechos.append(dict(empresa=emp, ruta=ruta, movimientos=len(mov), revisar=int(mov.revisar.sum()),
                           cuentas=len(cuentas[cuentas.movimientos > 0]), descuadres=len(descuadre)))
        avisar(f'  {emp:5s} {len(mov):5d} movimientos en {len(cuentas[cuentas.movimientos > 0])} cuentas | '
               f'revisar {int(mov.revisar.sum()):4d} | cuentas con descuadre de saldo: {len(descuadre)} -> {ruta}')
    return hechos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('libros', nargs='+')
    ap.add_argument('--mes', help='AAAA-MM (por defecto, el mes anterior)')
    ap.add_argument('--modelo', default=os.path.join(AQUI, 'modelo.pkl'))
    ap.add_argument('--salida', default='.')
    ap.add_argument('--umbral', type=float, default=0.9)
    a = ap.parse_args()
    anio, mes = map(int, a.mes.split('-')) if a.mes else mes_anterior()
    modelo = pickle.load(open(a.modelo, 'rb'))
    archivos = []
    for x in a.libros:
        archivos += sorted(glob.glob(os.path.join(x, '*.xlsx'))) if os.path.isdir(x) else [x]
    archivos = [f for f in archivos if not os.path.basename(f).startswith('~$')]
    clasificar(archivos, anio, mes, modelo, a.salida, a.umbral)


if __name__ == '__main__':
    main()
