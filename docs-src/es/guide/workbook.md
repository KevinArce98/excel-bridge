---
title: Edita un libro
description: Crea un libro con Workbook, o carga uno, cambia celdas, estilos y diseño, y guárdalo de nuevo.
group: Guía
groupOrder: 2
order: 7
---

# Edita un libro

`Workbook` crea un libro, o carga uno, lo edita y lo guarda de nuevo, sin reconstruir a mano los datos de las hojas. Reconstruye el archivo a partir de lo que modela la biblioteca, así que sirve para archivos que escribió esta biblioteca y para archivos simples de otras herramientas.

```ts title="workbook.ts"
import fs from 'node:fs';
import { Workbook } from 'excel-bridge';

const wb = Workbook.create();
wb.addSheet('Sales', [
  ['Product', 'Price', 'Qty'],
  ['Laptop', 999.99, 5],
]);
wb.setCellStyle('Sales', 0, 0, { bold: true, background: '#4472C4', color: '#FFFFFF' });
wb.setFreezePane('Sales', { row: 1 });
wb.setMetadata({ creator: 'My App', title: 'Sales Report' });
fs.writeFileSync('sales.xlsx', wb.toBuffer());

const existing = Workbook.fromBuffer(fs.readFileSync('sales.xlsx'));
existing.setCellValue('Sales', 1, 1, 899.99);
const blob = existing.toBlob();
```

Carga con `Workbook.fromBuffer(buffer)`, o con `await Workbook.fromFile(file)` en el navegador. Ambos aceptan los mismos [límites del lector](../errors/) que `ExcelReader`. Guarda con `toBuffer()` o `toBlob()`.

## Métodos

Las filas y las columnas empiezan en cero, y las hojas se identifican por nombre.

| Área | Métodos |
| --- | --- |
| Hojas | `addSheet`, `renameSheet`, `removeSheet`, `getSheetNames`, `getSheetData`, `getSheetState`, `setSheetState` |
| Celdas | `getCellValue`, `setCellValue`, `getCellStyle`, `setCellStyle` |
| Diseño | `setMergeCells`, `setFreezePane`, `setColumnWidths`, `setAutoWidth`, `setRowHeight`/`getRowHeight`, `setRowHidden`/`isRowHidden`, `setColumnHidden`/`isColumnHidden` |
| Reglas y vínculos | `addValidation`, `addConditionalFormat`, `setAutoFilter`, `getAutoFilter`, `removeAutoFilter`, `setHyperlink`, `getHyperlinks`, `removeHyperlink` |
| Documento | `getMetadata`, `setMetadata` |
| Salida | `toBuffer`, `toBlob` |

## Qué obtienes de un archivo cargado

`getCellValue` y `getSheetData` devuelven los mismos valores que escribes. Una fórmula cargada vuelve como `{ formula, result? }` con el valor guardado en el archivo, y un error cargado como `{ error }`. El texto sigue siendo texto, incluso cuando empieza con `=`.

> [!WARNING]
> El `result` de una fórmula es el valor que guardó el archivo. No cambia cuando editas las celdas que usa la fórmula, así que puede quedar desactualizado hasta que Excel recalcule al abrir. Asigna a la celda `{ formula }` sin `result` para descartarlo.

`getSheetData(name)` tiene huecos donde el archivo tiene filas en blanco: `length` es la última fila más uno, `for...of` produce `undefined` en un hueco y `JSON.stringify` escribe `null`. Una hoja cargada cuyo nombre Excel rechazaría, por ejemplo uno de más de 31 caracteres, se carga, pero guardar lanza un error hasta que le des un nombre válido con `renameSheet`.

## Qué conserva un ciclo de ida y vuelta

`Workbook.fromBuffer` y `fromFile` restauran:

- los datos en su fila y columna, los estilos, las celdas combinadas, los paneles inmovilizados y los anchos de columna
- los bordes por lado, las alturas de fila, las filas ocultas y las columnas ocultas
- las validaciones de datos, las reglas de formato condicional, los rangos de autofiltro, los hipervínculos y la visibilidad de las hojas
- las fórmulas con sus resultados guardados y los siete valores de error clásicos

## Qué cambia o descarta

- Imágenes, gráficos, comentarios, tablas de Excel, tablas dinámicas, nombres definidos, configuración de impresión, temas y macros.
- Los criterios de filtro y el estado de ordenación definidos en Excel. Las filas ocultas por un filtro siguen ocultas, y `removeAutoFilter` muestra las filas bajo el rango eliminado.
- Los vínculos que el escritor no acepta (todo excepto `http:`, `https:`, `mailto:` o una ubicación dentro del libro) y las validaciones de un tipo que el lector no conoce. `ExcelReader` sigue devolviendo los vínculos.
- Los mensajes de entrada y de error de la validación, el estilo de error (detener, advertencia, información) y los interruptores de "mostrar mensaje". Toda regla guardada se escribe para mostrar ambos mensajes y para rechazar las entradas incorrectas.
- El texto de error distinto de los siete errores clásicos, que se guarda como texto.
- Los formatos integrados de fecha y hora que dependen de la configuración regional, que se vuelven a escribir como códigos explícitos en-US. La fecha corta integrada (id 14) sigue siendo la fecha corta.
- Los colores de tema e indexados, y los formatos de número integrados como `0%`. Solo se leen los colores RGB, los formatos de número personalizados y los formatos integrados de fecha y hora, y los colores de borde de tema o indexados pasan a ser negros.
- Los bordes diagonales, el texto enriquecido (se aplana) y los paneles divididos (se descartan).
- Las reglas de formato condicional distintas de `cellIs`, `expression` y `colorScale`.
- Los valores predeterminados de la hoja (`<sheetFormatPr>`: altura de fila y ancho de columna predeterminados, niveles de esquema). Un grupo de esquema contraído sigue oculto, pero pierde sus botones.
- Los seguidores de una fórmula compartida, que se leen sin su fórmula.
- Una fórmula almacenada cuyo texto empieza con `=`. Excel no escribe una, pero un archivo hecho por 1.x a partir de una cadena como `=== Summary ===` la tiene. `Workbook` la guarda sin ese primer `=`, y una fórmula que es solo `=` hace que `toBuffer` lance un error.

Una columna guardada con ancho 0 sigue teniendo ancho 0 cuando la muestras con `setColumnHidden`. Define también su ancho.

> [!LIMIT]
> Si necesitas editar un archivo y conservar todo lo que no tocaste, como una plantilla con gráficos o macros, `Workbook` no es la herramienta adecuada: reconstruye el archivo a partir de su modelo. `openXlsx` y `saveXlsx` de hucre conservan las partes que no modelan.
