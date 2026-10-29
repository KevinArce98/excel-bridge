---
title: Rows as objects
description: Turn an array of objects into a sheet, and a sheet back into typed objects, with a header row or an explicit column list.
group: Guide
groupOrder: 2
order: 6
---

# Rows as objects

Most exports start as an array of objects and most imports end as one. Two functions cover each direction without hand-building cell arrays.

## Write objects

`objectsToSheet(rows, columns, options?)` returns an `ExcelData` you can hand to `ExcelWriter`. Each column names the property to read and, optionally, the header text, width, style and number format.

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

- The header row is always written, even when there are no rows.
- A column's `style` and `numberFormat` apply to its body cells, not to the header. Use `headerStyle` for the header.
- `undefined` and `null` give an empty cell, and properties that are not listed are ignored.
- Column widths replace `options.columnWidths`, so a column without a `width` gets Excel's default width and `autoWidth` has no effect on that sheet.
- `rows` can be any iterable. It is read once.

## Read objects

`sheetToObjects(sheet, options?)` reads the header row, then returns one object per data row.

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

| Case | Result |
| --- | --- |
| No `columns` | Keys are the header texts, in sheet order. |
| `columns` given | Only the listed columns, in the listed order, matched by `header` (default: the key). `column` picks a zero-based column instead. |
| Repeated header | `a`, `a_1`, `a_2`. |
| Blank header cell | Skipped. |
| Empty cell | `null`, so every object has the same keys. |
| Date-formatted cell | `Date`. |
| Error cell | `{ error: '#N/A' }` for the seven classic errors, otherwise the text. |
| Formula cell | Its cached value, or `null` if it has none. |
| Row with no value in any listed column | Skipped. |
| Listed column missing from the header | Throws `INVALID_FILE`, unless the column has `optional: true`. |

`headerRow` is the zero-based index of the header row and defaults to `0`. Pass `headerRow: null` when the sheet has no header, and give each column a `column` index.

> [!NOTE]
> This reads values only. It does not validate them. For schema-validated imports with error lists, read-excel-file is built for that.

## Stream objects

`objectsToStreamingSheet(rows, columns, options?)` does the same for the [streaming writer](../streaming/), from a sync or async iterable, without holding the rows in memory.

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
> Streaming columns take no `style` or `numberFormat`: the stream reads styles from a record fixed before the first row, so a body style would need one entry per row. Style the header with `headerStyle`.
