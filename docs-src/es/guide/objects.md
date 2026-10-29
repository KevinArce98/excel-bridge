---
title: Filas como objetos
description: Convierte un arreglo de objetos en una hoja, y una hoja de nuevo en objetos con tipo, con una fila de encabezado o una lista explícita de columnas.
group: Guía
groupOrder: 2
order: 6
---

# Filas como objetos

La mayoría de las exportaciones empiezan como un arreglo de objetos y la mayoría de las importaciones terminan como uno. Dos funciones cubren cada dirección sin construir a mano arreglos de celdas.

## Escribir objetos

`objectsToSheet(rows, columns, options?)` devuelve un `ExcelData` que puedes entregar a `ExcelWriter`. Cada columna nombra la propiedad que se lee y, de forma opcional, el texto del encabezado, el ancho, el estilo y el formato de número.

```ts title="objects-to-sheet.ts"
import { ExcelWriter, objectsToSheet } from 'excel-bridge';

interface Order {
  id: number;
  customer: string;
  total: number;
  placed: Date;
}

const orders: Order[] = [
  { id: 1, customer: 'Ana', total: 120.5, placed: new Date(2024, 0, 15) },
  { id: 2, customer: 'Ben', total: 89, placed: new Date(2024, 0, 16) },
];

const sheet = objectsToSheet<Order>(
  orders,
  [
    { key: 'id', header: 'Order', width: 10 },
    { key: 'customer', header: 'Customer', width: 24 },
    { key: 'total', header: 'Total', width: 14, numberFormat: '#,##0.00' },
    { key: 'placed', header: 'Placed', width: 14, numberFormat: 'yyyy-mm-dd' },
  ],
  {
    name: 'Orders',
    freezePane: { row: 1 },
    headerStyle: { bold: true, background: '#D9E1F2' },
  }
);

const bytes = new ExcelWriter().createWorkbookBuffer([sheet]);
```

- La fila de encabezado siempre se escribe, incluso cuando no hay filas.
- El `style` y el `numberFormat` de una columna se aplican a sus celdas del cuerpo, no al encabezado. Usa `headerStyle` para el encabezado.
- `undefined` y `null` producen una celda vacía, y las propiedades que no se enumeran se ignoran.
- Los anchos de columna reemplazan a `options.columnWidths`, así que una columna sin `width` recibe el ancho predeterminado de Excel y `autoWidth` no tiene efecto en esa hoja.
- `rows` puede ser cualquier iterable. Se lee una sola vez.

## Leer objetos

`sheetToObjects(sheet, options?)` lee la fila de encabezado y luego devuelve un objeto por cada fila de datos.

```ts title="sheet-to-objects.ts"
import fs from 'node:fs';
import { ExcelReader, sheetToObjects } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('data.xlsx'));

const rows = sheetToObjects(workbook.sheets[0]);

interface Line {
  name: string;
  qty: number | null;
}

const typed = sheetToObjects<Line>(workbook.sheets[0], {
  columns: [
    { key: 'name', header: 'Name' },
    { key: 'qty', header: 'Qty', parse: value => (typeof value === 'number' ? value : null) },
  ],
});
```

| Caso | Resultado |
| --- | --- |
| Sin `columns` | Las claves son los textos del encabezado, en el orden de la hoja. |
| Con `columns` | Solo las columnas enumeradas, en el orden indicado, emparejadas por `header` (por defecto: la clave). `column` elige en su lugar una columna base cero. |
| Encabezado repetido | `a`, `a_1`, `a_2`. |
| Celda de encabezado en blanco | Se omite. |
| Celda vacía | `null`, de modo que todos los objetos tienen las mismas claves. |
| Celda con formato de fecha | `Date`. |
| Celda de error | `{ error: '#N/A' }` para los siete errores clásicos; en otro caso, el texto. |
| Celda con fórmula | Su valor en caché, o `null` si no tiene. |
| Fila sin valor en ninguna columna enumerada | Se omite. |
| Columna enumerada que falta en el encabezado | Lanza `INVALID_FILE`, salvo que la columna tenga `optional: true`. |

`headerRow` es el índice base cero de la fila de encabezado y su valor predeterminado es `0`. Pasa `headerRow: null` cuando la hoja no tiene encabezado, y dale a cada columna un índice `column`.

> [!NOTE]
> Esto solo lee valores. No los valida. Para importaciones validadas con un esquema y con listas de errores, read-excel-file está pensado para eso.

## Objetos en streaming

`objectsToStreamingSheet(rows, columns, options?)` hace lo mismo para el [escritor de streaming](../streaming/), a partir de un iterable síncrono o asíncrono, sin mantener las filas en memoria.

```ts title="objects-to-streaming-sheet.ts"
import { createExcelWorkbookStream, objectsToStreamingSheet, streamToBuffer } from 'excel-bridge';

async function* orders() {
  for (let id = 1; id <= 5; id++) yield { id, customer: `Customer ${id}`, total: id * 10 };
}

const bytes = await streamToBuffer(
  createExcelWorkbookStream([
    objectsToStreamingSheet(
      orders(),
      [
        { key: 'id', header: 'Order', width: 10 },
        { key: 'customer', header: 'Customer', width: 24 },
        { key: 'total', header: 'Total', width: 14 },
      ],
      { name: 'Orders', freezePane: { row: 1 }, headerStyle: { bold: true } }
    ),
  ])
);
```

> [!LIMIT]
> Las columnas de streaming no aceptan `style` ni `numberFormat`: el stream lee los estilos de un registro fijado antes de la primera fila, así que un estilo del cuerpo necesitaría una entrada por fila. Dale estilo al encabezado con `headerStyle`.
