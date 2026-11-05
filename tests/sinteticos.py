"""Datos inventados con la misma estructura que los reales. Ningún nombre, cuenta o monto es real."""
from __future__ import annotations

import datetime as dt
import random

import openpyxl
import pandas as pd

D = dt.datetime
COLUMNAS = ['CH/Nº', 'FECHA', 'CONCEPTO', 'DEBE', 'HABER', 'SALDO', 'N OPERACION']


def hoja_libro(wb, nombre: str, filas: list[tuple], saldo: float = 1000.0,
               cabecera: str = 'CTA.CTE. M.E. N 000-000') -> None:
    """Agrega una cuenta. Cada fila es (fecha, concepto, debe, haber); el saldo corrido sigue el orden de las filas."""
    ws = wb.create_sheet(nombre)
    ws.append(['EMPRESA DE PRUEBA'])
    ws.append([cabecera])
    ws.append([])
    ws.append(COLUMNAS)
    for f, c, d, h in filas:
        saldo += (d or 0) - (h or 0)
        ws.append([None, f, c, d, h, round(saldo, 2), None])


def libro_basico(ruta: str) -> str:
    """Libro de setiembre con dos cuentas que cuadran: dólares con ITF en la misma fila y soles."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    hoja_libro(wb, 'BBVA$', [
        (D(2026, 8, 31), 'CIERRE AL 31/08', None, None),
        (D(2026, 9, 1), 'DEPOSITO CLIENTE ALFA', 2000.0, 0.10),
        (D(2026, 9, 2), 'PAGO PROVEEDOR BETA', None, 500.0),
        (D(2026, 9, 3), 'COMISION MANTENIMIENTO', None, 15.0),
        (D(2026, 9, 4), 'OVERNIGHT DEL 3 AL 4 TASA 2.5%', 100.0, None),
        (D(2026, 9, 5), 'BCR AL BCP', None, 300.0),
        (D(2026, 9, 6), 'VENTA DOLARES TC 3.40', None, 200.0),
    ])
    hoja_libro(wb, 'BCP S', [
        (D(2026, 8, 31), 'CIERRE AL 31/08', None, None),
        ('5-Set-26', 'DEPOSITO CLIENTE ALFA', 900.0, None),
        ('12-Set-26', 'PAGO PLANILLA GAMA', None, 400.0),
    ], cabecera='CTA.CTE. M.N. N 000-111')
    wb.create_sheet('SALDOS').append(['Hoja que no es una cuenta'])
    wb.save(ruta)
    return ruta


def historico(empresas: tuple[str, ...] = ('ICP', 'SBSA'), repeticiones: int = 12) -> pd.DataFrame:
    """Histórico mínimo con conceptos inventados para entrenar el modelo."""
    azar = random.Random(7)
    plantillas = [
        ('DEPOSITO CLIENTE ALFA', 'INGRESO', 'DEPOSITOS', 1),
        ('DEPOSITO CLIENTE DELTA', 'INGRESO', 'DEPOSITOS', 1),
        ('OVERNIGHT DEL 3 AL 4 TASA 2.5%', 'INGRESO', 'ABONO TRANSF. CTAS', 2),
        ('PAGO PROVEEDOR BETA', 'EGRESO', 'PAGO PROVEEDORES', 3),
        ('PAGO PROVEEDOR EPSILON', 'EGRESO', 'PAGO PROVEEDORES', 3),
        ('PAGO PLANILLA GAMA', 'EGRESO', 'PAGO HABERES', 4),
        ('COMISION MANTENIMIENTO', 'EGRESO', 'GASTOS BANCARIOS', 5),
        ('BCR AL BCP', 'EGRESO', 'CARGO TRANSF. CTAS', 6),
        ('VENTA DOLARES TC 3.40', 'EGRESO', 'CARGO TRANSF. CTAS', 6),
        ('ITF', 'EGRESO', 'ITF', 7),
    ]
    filas = []
    for e in empresas:
        for _ in range(repeticiones):
            for concepto, tipo, cat, orden in plantillas:
                filas.append(dict(itf_misma_fila=0, empresa=e, periodo='2026-08', hoja='BBVA$', banco='BBVA',
                                  moneda='USD', fecha='2026-08-10', concepto=concepto, tipo=tipo,
                                  monto=round(azar.uniform(10, 5000), 2), categoria=cat, orden=orden,
                                  archivo=f'FLUJO {e} AGOSTO 2026.xlsx'))
    # SBSA dejó un ITF de la misma fila dentro del movimiento: la regla debe mandarlo igual a ITF
    filas.append(dict(itf_misma_fila=1, empresa='SBSA', periodo='2026-08', hoja='BBVA$', banco='BBVA', moneda='USD',
                      fecha='2026-08-10', concepto='DEPOSITO CLIENTE ALFA', tipo='EGRESO', monto=0.1,
                      categoria='DEPOSITOS', orden=1, archivo='FLUJO SBSA AGOSTO 2026.xlsx'))
    return pd.DataFrame(filas)
