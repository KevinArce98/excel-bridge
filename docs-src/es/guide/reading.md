---
title: Lee un libro
description: Analiza un .xlsx desde un Buffer o un File en celdas con tipo, estilos y diseño, y maneja las filas que faltan.
group: Guía
groupOrder: 2
order: 5
---

# Lee un libro

`ExcelReader` convierte un `.xlsx` en objetos simples: cada hoja con sus celdas con tipo, sus estilos y su diseño. La lectura es síncrona para los bytes que ya tienes y asíncrona para un `File` del navegador.

```ts title="read.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const reader = new ExcelReader();

const workbook = reader.parseFromBuffer(fs.readFileSync('data.xlsx'));

const [sheet] = workbook.sheets;
console.log(sheet.name, sheet.data.length);
```

En el navegador, pasa el `File` de un `<input type="file">` a `await reader.parseFromFile(file)`. `ExcelBridge.read(buffer)` y `parseExcel(buffer)` hacen lo mismo que `parseFromBuffer`.

## Las celdas se tipan con `type`

Cada celda tiene una de seis formas, que se distinguen por `type`. Comprueba `type` y TypeScript acota `value` por ti.

```ts title="narrow.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('data.xlsx'));

workbook.sheets[0].data.forEach(row => {
  row.forEach(cell => {
    switch (cell.type) {
      case 'number':
        console.log(cell.coordinate, cell.value.toFixed(2));
        break;
      case 'date':
        console.log(cell.coordinate, cell.value.toISOString());
        break;
      case 'error':
        console.log(cell.coordinate, 'error', cell.value);
        break;
      case 'empty':
        break;
      default:
        console.log(cell.coordinate, cell.value);
    }
  });
});
```

| `type` | `value` | Notas |
| --- | --- | --- |
| `'string'` | `string` | Los escapes como `_x000D_` se decodifican. Una celda `t="d"` que no es una de las formas de fecha ISO de abajo se lee como su texto. |
| `'number'` | `number` | |
| `'boolean'` | `boolean` | |
| `'date'` | `Date` | Una celda cuyo formato de número es de fecha, o una celda `t="d"` escrita como `YYYY-MM-DD` o `YYYY-MM-DDTHH:MM:SS` con fracción opcional y una `Z` final. Hora local de reloj: se ignora la `Z`. Una fecha con desfase como `+05:30`, o con un espacio en lugar de `T`, se lee como texto, igual que la muestra Excel. Se leen los dos sistemas de fechas (1900 y 1904). |
| `'error'` | `string` | El texto del error, como `#DIV/0!`. |
| `'empty'` | `null` | Una celda con estilo pero sin valor, o relleno de un hueco entre celdas de una fila. |

Cada celda también tiene `coordinate` (`"B2"`), `rowIndex` y `columnIndex`, ambos base cero. Una celda con fórmula guarda su fórmula en `formula`, sin el `=` inicial, y su valor en caché en `value`. Una fórmula sin valor en caché es una celda `'empty'` con una `formula`.

## Las filas se indexan por posición

`sheet.data` se indexa por la posición de la fila en la hoja, así que `sheet.data[4]` es la fila 5. Las filas que no están en el archivo son huecos en el arreglo.

```ts title="holes.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('data.xlsx'));
const { data } = workbook.sheets[0];

data.forEach((row, index) => console.log(index + 1, row.length));

for (const row of data) {
  if (!row) continue;
  console.log(row.map(cell => cell.value).join(' | '));
}
```

> [!NOTE]
> `forEach`, `map` y `Object.values` omiten los huecos. Un bucle `for...of` visita cada índice y produce `undefined` en un hueco, así que añade una cláusula de guarda. `data.length` es la última fila más uno.

## Estilos, combinaciones y diseño

Además de las celdas, un `ParsedSheet` expone lo que el archivo dice sobre la hoja.

```ts title="sheet-details.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const { sheets, metadata } = new ExcelReader().parseFromBuffer(fs.readFileSync('data.xlsx'));
const sheet = sheets[0];

sheet.name;
sheet.state;
sheet.styles;
sheet.mergeCells;
sheet.freezePane;
sheet.columnWidths;
sheet.rowHeights;
sheet.hiddenRows;
sheet.hiddenColumns;
sheet.validations;
sheet.conditionalFormats;
sheet.autoFilter;
sheet.hyperlinks;
metadata.creator;
```

- **`state`** es `'hidden'` o `'veryHidden'` para una hoja oculta, y `undefined` en caso contrario.
- **`styles`** asocia `"row-col"` con un `CellStyle`, el mismo tipo que recibe el escritor. En un estilo leído, `border` es `true` para un recuadro negro delgado en los cuatro lados y, en cualquier otro caso, un objeto con una entrada `{ style, color? }` por cada lado que tiene línea. Solo se leen los colores RGB y los formatos de número personalizados, además de los formatos integrados de fecha y hora.
- **`rowHeights`, `hiddenRows` y `hiddenColumns`** aparecen solo en los archivos que los tienen.
- **`hyperlinks`** se devuelven tal como están guardados, sea cual sea su esquema. Comprueba `url` antes de mostrarlo como un `<a href>`.

## Límites para archivos que no creaste tú

`new ExcelReader({ maxCells, maxPartBytes, maxTotalBytes, maxSheets })` define límites para archivos que no creaste tú. Los límites de tamaño y `maxSheets` se comprueban antes de descomprimir las hojas. `maxCells` se comprueba mientras se analiza una hoja, después de descomprimir su XML, así que usa `maxPartBytes` para acotar la memoria. Consulta [Maneja errores y límites](../errors/).

## Filas como objetos

Para leer una tabla como un arreglo de objetos en lugar de celdas, consulta [Filas como objetos](../objects/).
