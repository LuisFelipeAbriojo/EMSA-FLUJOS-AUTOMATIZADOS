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
- Separa el ITF que viene en la misma fila (BBVA y otros) y verifica que los saldos cuadren por cuenta.

| Archivo | Qué hace |
|---|---|
| `index.html` | La app web (generada; no editar a mano) |
| `web/plantilla.html` | Interfaz de la app |
| `app_web.py` | Puente entre la página y el clasificador |
| `leer_libro.py` | Lee el libro bancos y extrae los movimientos y saldos del mes |
| `modelo.py` | Clasificador (memoria + regresión logística) |
| `salida.py` | Escribe el Excel por empresa |
| `clasificar_mes.py`, `actualizar.py`, `entrenar.py` | Versión de línea de comandos |
| `extraer_historico.py` | Arma `historico.csv` desde RESUMEN FLUJO cerrados |
| `construir_web.py` | Regenera `index.html` |

## Para desarrollo
- Si cambias algún `.py`, regenera la app: `python construir_web.py`, luego commit y push.
- Versión de escritorio (requiere Python 3.10+): ver [LEEME.md](LEEME.md).
