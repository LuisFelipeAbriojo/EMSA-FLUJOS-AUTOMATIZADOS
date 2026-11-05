# Flujos de caja

Del libro bancos al **RESUMEN FLUJO** de cada empresa: clasifica los movimientos del mes, arma el flujo con su formato de siempre
y aprende de los flujos ya corregidos.

**👉 App: https://luisfelipeabriojo.github.io/EMSA-FLUJOS-AUTOMATIZADOS/**

No hay que instalar nada. También funciona abriendo el `index.html` con doble clic desde el disco compartido.

## Uso mensual
Ejemplo: en noviembre se cierra octubre. Usa Chrome o Edge. La primera carga tarda unos segundos porque descarga el motor.

### 1. Modelo
Carga `modelo_web.pkl` (instantáneo) o el `historico.csv` vigente (entrena en unos minutos). Los dos están en el disco compartido.

### 2. Clasificar
1. Arrastra los libros bancos del mes, un `.xlsx` por empresa con la empresa en el nombre: `ICP OCT 2026.xlsx`.
2. Elige el mes y pulsa **Clasificar**. Sale un Excel clasificado por empresa.
3. **Revisa la hoja MOVIMIENTOS** de cada uno. Las filas amarillas tienen confianza baja. La columna CATEGORIA tiene lista desplegable.

### 3. Armar flujo
1. Arrastra, por empresa, el Excel clasificado ya revisado y el RESUMEN FLUJO del mes anterior.
2. Elige el mes, escribe el **TC de fin de mes** y pulsa **Armar flujo**.
3. Descarga el RESUMEN FLUJO del mes y sus **notas para la Jefa de Finanzas**: controles, alertas, variaciones y supuestos.
4. Revisa cada flujo en Excel y **guárdalo**. Un flujo que nadie volvió a guardar no puede entrar al histórico.

### 4. Actualizar histórico
Cuando los flujos del mes estén corregidos y entregados:
1. Arrastra un `.zip` con esos flujos y el `historico.csv` vigente.
2. Pulsa **Actualizar histórico**. Si una empresa y mes ya estaban, se reemplazan.
3. Descarga el histórico nuevo (`historico_AAAA-MM-DD.csv`), el reporte de lo que entró y el modelo reentrenado.
4. Guárdalos en el disco compartido, **sin borrar las versiones anteriores**.

## Privacidad
Todo se procesa **dentro del navegador** con [Pyodide](https://pyodide.org). Los libros, los flujos y el histórico **no se envían a ningún servidor**.
Internet solo se usa para descargar el motor desde `cdn.jsdelivr.net`.

Este repositorio es público y contiene solo código. Ningún dato de las empresas puede subirse: lo impiden `.gitignore`,
un hook de pre-commit y una revisión en cada integración continua.

## Resultados del clasificador
Validado con agosto de 2026 contra la clasificación del equipo: **99,1 %** de acierto en 12.623 movimientos.
Con el umbral de 90 % se revisa cerca del 2,5 % de las filas, y lo que pasa sin revisar acierta 99,6 %.

## Limitaciones conocidas
- **Revisión obligatoria:** el armado es un borrador con controles, no reemplaza la revisión del flujo.
- **Hojas Over:** las filas disponibles son fijas. Si hay más overnights que filas, el que no entra se avisa en las notas y hay que agregarlo a mano.
- **Ajuste de diferencia de cambio:** no se aplica solo. Se sugiere en las notas.
- **Secciones nuevas:** no se crean solas. Sus movimientos se listan en las notas y hay que agregar la sección a mano.
- **"Sin revisar":** el histórico detecta que nadie guardó el flujo, no qué tan a fondo se revisó.
- **Impresión:** el flujo armado no conserva la configuración de impresión.

## Documentación
- [Arquitectura](docs/arquitectura.md)
- [Inventario de ajustes mensuales](docs/ajustes-mensuales.md)
- [Cambios por versión](CHANGELOG.md)

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

- La línea de comandos es para desarrollo: `python -m flujos --help`.
- Las pruebas con datos reales corren si `FLUJOS_DATOS` apunta a la carpeta privada; si no, se saltan.
- Cada push a `main` que pasa las pruebas publica la app en GitHub Pages.
