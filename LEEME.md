# Clasificador de flujos de caja (corre en tu laptop, sin internet)

## Instalar (una sola vez)
1. Instalar Python 3.10 o superior (python.org, marcar "Add Python to PATH").
2. En una terminal, dentro de esta carpeta:  pip install pandas openpyxl scikit-learn scipy

## Uso mensual (ej.: en noviembre se procesa octubre)
1. Copiar los libros bancos de cada empresa (ECSA OCT 2026.xlsx, ICP ..., etc.) a la carpeta LIBROS.
2. Doble clic en CLASIFICAR_MES.bat  (o:  python clasificar_mes.py LIBROS --salida SALIDA)
   - Toma automáticamente el mes anterior al actual. Para otro mes:  --mes 2026-09
   - La primera vez entrena el modelo (1-2 minutos).
3. En SALIDA queda un Excel por empresa:
   - RESUMEN: saldos por cuenta (rojo = el libro no cuadra) y totales por categoría.
   - MOVIMIENTOS: todo el mes; filas amarillas = revisar. La columna CATEGORIA tiene lista desplegable.
   - Una hoja por cuenta, agrupada por categorías como en los RESUMEN FLUJO.
4. Después de revisar y corregir en MOVIMIENTOS:
   python actualizar.py "SALIDA\ICP FLUJO OCTUBRE 2026 CLASIFICADO.xlsx"
   python entrenar.py historico.csv
   (rehace las hojas por cuenta y el modelo aprende de las correcciones)

## Archivos
- leer_libro.py: lee el libro bancos y se queda solo con el mes (soporta fechas escritas "1-Set-26").
- modelo.py: memoria de conceptos ya vistos + modelo de texto; solo propone categorías que esa empresa usa.
- comun.py: unifica variantes de nombre de categoría ("CARGO TRASNF CTAS" -> "CARGO TRANSF. CTAS").
- extraer_historico.py: arma historico.csv a partir de RESUMEN FLUJO ya cerrados (para agregar más meses).
- historico.csv: julio y agosto 2026 ya categorizados (25.698 movimientos).

## Versión web (sin instalar Python)
`index.html` es la misma herramienta corriendo dentro del navegador (Pyodide): no hace falta instalar nada
y los archivos no salen de la computadora. Se puede abrir con doble clic desde el disco compartido o publicar
en GitHub Pages. Pasos: 1) cargar `historico.csv` o `modelo_web.pkl`, 2) arrastrar los libros bancos del mes,
3) (opcional) subir los Excel revisados para actualizar el histórico.
- Si cambias algún .py, vuelve a generar la página: `python construir_web.py`
- El repositorio NO debe incluir datos (historico.csv, libros, salidas, modelos): ver `.gitignore`.
