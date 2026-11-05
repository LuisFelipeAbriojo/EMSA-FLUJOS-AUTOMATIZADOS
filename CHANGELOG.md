# Cambios

## 0.3.1 (sin publicar)
### Corregido
- El saldo inicial del mes nuevo en TOTAL siempre sigue al saldo final del mes anterior, aunque la columna ya trajera otra fórmula.
- Las fórmulas que apuntan a otro libro quedan con su valor guardado, en vez de salir como `#REF!`.

### Agregado
- Control de líneas de la hoja mensual que no llegan a TOTAL.
- Aviso de diferencia entre el saldo final de TOTAL y el de la hoja mensual.

## 0.3.0
### Agregado
- Paso **Actualizar histórico**: recibe un `.zip` con flujos corregidos y el `historico.csv` vigente, y entrega el histórico con fecha, un reporte y el modelo reentrenado.
- Reemplazo por empresa y periodo, sin tocar lo demás.
- Rechazo de Excel clasificados, notas, proyectados, archivos dañados, duplicados, flujos con fechas de otro mes y flujos armados que nadie volvió a guardar.
- Comandos `armar`, `historico` y `rehacer` en la línea de comandos.
- Progreso visible en los pasos largos.
- `docs/` con la arquitectura y el inventario de ajustes mensuales.

### Cambiado
- La interfaz queda en cuatro pasos: Modelo, Clasificar, Armar flujo y Actualizar histórico.
- "Enseñarle al modelo" ya no existe: el histórico solo aprende de flujos entregados. Rehacer las hojas por cuenta queda como opción del paso 2.
- Las filas de flujos cerrados con la fecha escrita como texto ahora entran al histórico.

### Quitado
- `index.html` de la raíz y `LEEME.md`: la app se construye y publica sola.

## 0.2.0
### Agregado
- Paso **Armar flujo**: del Excel clasificado y el flujo anterior sale el RESUMEN FLUJO del mes, con validaciones y notas para la Jefa de Finanzas.
- Evaluador de fórmulas propio, para validar y guardar valores calculados sin Excel.
- Llenado de hojas Over y FM, y cálculo de intereses overnight como abono menos cargo.
- Lectura del formato de libro de ELSA.

### Cambiado
- El proyecto pasa a ser un paquete `src/flujos` con configuración TOML, logging, ruff y pytest.
- La interfaz se separa en HTML, CSS y JS.
- GitHub Actions valida privacidad, estilo y pruebas, y publica la app.

### Seguridad
- Candado de privacidad: `.gitignore` estricto, hook de pre-commit y correo noreply en los commits.

## 0.1.0
- Clasificador en el navegador con Pyodide, publicado en GitHub Pages.
- ITF de la misma fila separado a su categoría.
- Transferencias ordenadas por tipo de operación y fecha.
- Fechas mal digitadas contadas según su ubicación en el libro.
