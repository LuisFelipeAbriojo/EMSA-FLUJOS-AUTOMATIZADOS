# Clasificador de flujos de caja

Clasifica automáticamente los movimientos bancarios del mes (DEPOSITOS, ITF, PAGO PROVEEDORES, TRANSF. ENTRE CIAS, etc.)
a partir de los **libros bancos** de cada empresa y arma un Excel por empresa con el formato del RESUMEN FLUJO.

**👉 App: https://luisfelipeabriojo.github.io/EMSA-FLUJOS-AUTOMATIZADOS/**

## Cómo se usa (sin instalar nada)
1. Abre la app en Chrome o Edge. La primera carga tarda ~15 s (descarga el motor).
2. **Modelo:** carga `historico.csv` (entrena en ~1 min) o `modelo_web.pkl` si ya lo tienes (instantáneo). Ambos están en el disco compartido.
3. **Libros bancos:** arrastra los `.xlsx` del mes (uno por empresa, con la empresa en el nombre: `ICP OCT 2026.xlsx`),
   elige el mes y pulsa *Clasificar*. Descarga un Excel por empresa o un .zip con todos.
4. **Revisa** en cada Excel la hoja MOVIMIENTOS: las filas amarillas tienen confianza baja. La columna CATEGORIA tiene lista desplegable.
5. **(Opcional) Aprende de las correcciones:** sube los Excel revisados en el paso 3 de la app; descarga el `historico.csv`
   actualizado y reemplaza el del disco compartido.

## Privacidad
Todo se procesa **dentro del navegador** (Python vía [Pyodide](https://pyodide.org)). Los libros, el histórico y los resultados
**no se envían a ningún servidor**. Internet solo se usa para descargar el motor (cdn.jsdelivr.net y pypi.org).
Este repositorio contiene solo código: ningún dato de las empresas (ver `.gitignore`).

## Resultados
Validado con agosto 2026 contra la clasificación del equipo: **99,1 %** de acierto en 12.623 movimientos.
Con el umbral de 90 % se revisa ~2,5 % de las filas, y el resto acierta 99,6 %.

## Cómo funciona
- **Memoria:** si un concepto (misma empresa y tipo) siempre tuvo la misma categoría, se repite.
- **Modelo:** regresión logística sobre el texto del concepto, códigos, empresa, banco, moneda y rango de monto.
  Solo propone categorías que esa empresa ya usó.
- El ITF que viene en la misma fila (BBVA y otros) se separa y va siempre a la categoría ITF.
- En ABONO / CARGO TRANSF. CTAS ordena por tipo de operación (overnight → BCR → cambio de moneda → fondos mutuos → otros)
  y luego por fecha. El orden se cambia en `TIPOS_OPERACION` de `comun.py`.
- Verifica que los saldos cuadren por cuenta. Si una fila o un bloque tiene un mes distinto al de las filas de arriba y abajo
  (fecha mal digitada), se cuenta según su ubicación en el libro, se marca en amarillo y se lista en RESUMEN → FECHAS SOSPECHOSAS.

## Para desarrollo
Requiere Python 3.11 o superior. Los datos reales viven fuera de este repositorio.

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -e ".[dev]"
git config core.hooksPath .githooks
.venv/Scripts/python -m pytest
.venv/Scripts/ruff check src tests scripts
.venv/Scripts/python scripts/construir_web.py
```

- `src/flujos/`: `lectura/` (libros y flujos), `clasificacion/` (modelo y reglas), `reportes/` (Excel), `historico/`, `app/` (puente con la web), `cli.py`.
- `web/`: la interfaz (HTML, CSS y JS). `scripts/construir_web.py` genera `dist/index.html` con el paquete incrustado.
- Línea de comandos: `python -m flujos --help`.
- Pruebas con datos reales: define `FLUJOS_DATOS` con la carpeta privada; si no existe, se saltan.
