# Inventario de ajustes mensuales

Qué cambia de un mes a otro en un RESUMEN FLUJO, según la comparación de dos meses seguidos de las 9 empresas. Este documento no incluye datos de las empresas.

## Automatizado en todas las empresas
| Hoja | Ajuste |
|---|---|
| Hoja mensual | Nombre de la hoja según el patrón de cada empresa (`AGOSTO`, `AGO` o `AGO-26`). |
| Hoja mensual | El TC del mes pasa a ser el anterior, y entra el nuevo. |
| Hoja mensual | Los vínculos apuntan al subtotal nuevo de cada sección. |
| TOTAL | La columna del mes anterior pasa de fórmula a valor. |
| TOTAL | La columna del mes nuevo se vincula a la hoja mensual. |
| TOTAL | El acumulado del año y las variaciones incluyen el mes nuevo. |
| Bancos | El saldo inicial es el saldo final anterior, como valor. |
| Bancos | Los bloques de movimientos se reemplazan por sección, con sus subtotales. |

## Automatizado con reglas propias
- **Intereses overnight:** interés = abono − cargo. El cargo puede estar en el mes anterior, y el libro trae esa historia.
- **Hojas Over:** registran solo los overnights que cruzan de un mes a otro. Los que vuelven usan hoy el TC anterior, para que neteen contra lo que salió.
- **Hojas FM:** existen solo en ECSA, ICP y SBSA. Se llenan desde el libro.

Hay tres formas de registrar el interés overnight, según la hoja:
1. El abono va bruto y una celda en HABER compensa el interés.
2. El abono va solo con el capital, y el interés va aparte.
3. El interés va en HABER de cada abono, y la sección INTERESES lleva solo el total.

## Solo se avisa en las notas
- **Ajustes a mano en TRANSF. ENTRE CTAS:** se quitan al armar y se reportan. La diferencia de cambio se sugiere, pero no se aplica.
- **Secciones nuevas:** sus movimientos se listan en las notas. Si la sección ya existe sin vínculo, el armado crea el vínculo y lo reporta.
- **Errores que ya venían en el archivo:** como `#REF!` o columnas antiguas vinculadas al mes. Se reportan sin tocarlos.
- **Vínculos a una fila de movimiento:** no se pueden arrastrar solos.

## Diferencias por empresa
| Empresa | TC (celdas) | FM | Observación |
|---|---|---|---|
| ECNA | E3:F3 | No | |
| ECSA | F3:G3 | Sí | |
| ELSA | N5:O5 | No | Tiene hojas ocultas y una hoja `Resumen`. Su libro usa otro nombre de columna y fechas con puntos. |
| ESSA | O3:O4 | No | |
| ICP | F3:G3 | Sí | |
| ONE | H3:H4 | No | |
| REVO | R1:R2 | No | Tiene una hoja oculta `obras`. |
| SBSA | R3:S3 | Sí | Algunas hojas tienen dos cuadros lado a lado. |
| T1 | G3:G4 | No | |

El orden de TC actual y anterior cambia entre empresas: el armado lo deduce de las fórmulas. Las filas de cada hoja Over están en `src/flujos/recursos/config.toml`.

## Qué se pierde al guardar con openpyxl
Medido con un guardado sin cambios en las 9 empresas. Los archivos no tienen gráficos, tablas dinámicas, tablas ni formato condicional.

| Elemento | Qué pasa |
|---|---|
| Estilos, anchos visibles, combinadas, paneles y hojas ocultas | Se conservan. |
| Valores calculados | openpyxl los pierde, pero `guardar.py` los vuelve a escribir. |
| Orientación, márgenes, escala, DPI y áreas de impresión | Se conservan. Medido en 124 hojas. |
| Ajustes del controlador de la impresora | Se pierden. Son binarios propios de cada impresora, y no se reinyectan. |
| Nombres ocultos de filtros antiguos | Se pierden, sin efecto visible. |
| Cuadros de texto vacíos de ELSA | Se pierden, sin efecto visible. |
| Comentarios | Se reescriben, y su tamaño puede cambiar. |

## Pendiente de confirmar con Finanzas
- Regla exacta de la hoja Over y del TC de los overnights que vuelven.
- Si la diferencia de cambio debe aplicarse sola, y qué son los ajustes que no calzan con esa regla.
- Si las secciones nuevas deben crearse solas.
- Reglas de las hojas FM deducidas de los archivos.
- Cómo identificar cada mes las líneas vinculadas a una fila de movimiento.
