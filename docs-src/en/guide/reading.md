---
title: Read a workbook
description: Parse an .xlsx from a Buffer or a File into typed cells, styles and layout, and handle rows that are missing.
group: Guide
groupOrder: 2
order: 5
---

# Read a workbook

`ExcelReader` turns an `.xlsx` into plain objects: every sheet with its typed cells, its styles and its layout. Reading is synchronous for bytes you already have, and asynchronous for a browser `File`.

```ts title="read.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const reader = new ExcelReader();

const workbook = reader.parseFromBuffer(fs.readFileSync('data.xlsx'));

const [sheet] = workbook.sheets;
console.log(sheet.name, sheet.data.length);
```

In the browser, pass the `File` from an `<input type="file">` to `await reader.parseFromFile(file)`. `ExcelBridge.read(buffer)` and `parseExcel(buffer)` do the same as `parseFromBuffer`.

## Cells are typed by `type`

Each cell is one of six shapes, told apart by `type`. Check `type` and TypeScript narrows `value` for you.

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

| `type` | `value` | Notes |
| --- | --- | --- |
| `'string'` | `string` | Escapes such as `_x000D_` are decoded. A `t="d"` cell that is not one of the ISO date forms below is read as its text. |
| `'number'` | `number` | |
| `'boolean'` | `boolean` | |
| `'date'` | `Date` | A cell whose number format is a date, or a `t="d"` cell written as `YYYY-MM-DD` or `YYYY-MM-DDTHH:MM:SS` with an optional fraction and a trailing `Z`. Local wall-clock time: the `Z` is ignored. A date with an offset such as `+05:30`, or a space instead of `T`, is read as text, as Excel shows it. Both date systems (1900 and 1904) are read. |
| `'error'` | `string` | The error text, such as `#DIV/0!`. |
| `'empty'` | `null` | A cell with a style but no value, or padding for a gap between cells in a row. |

Every cell also has `coordinate` (`"B2"`), `rowIndex` and `columnIndex`, both zero-based. A formula cell carries its formula in `formula`, without the leading `=`, and its cached value in `value`. A formula without a cached value is an `'empty'` cell with a `formula`.

## Rows are indexed by position

`sheet.data` is indexed by the row's position in the sheet, so `sheet.data[4]` is row 5. Rows that are not in the file are holes in the array.

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
> `forEach`, `map` and `Object.values` skip holes. A `for...of` loop visits every index and yields `undefined` for a hole, so guard it. `data.length` is the last row plus one.

## Styles, merges and layout

Besides the cells, a `ParsedSheet` exposes what the file says about the sheet.

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

- **`state`** is `'hidden'` or `'veryHidden'` for a hidden sheet, and `undefined` otherwise.
- **`styles`** maps `"row-col"` to a `CellStyle`, the same type the writer takes. In a read style, `border` is `true` for a thin black box on all four sides and otherwise an object with one `{ style, color? }` entry for each side that has a line. Only RGB colours and custom number formats are read, plus the built-in date and time formats.
- **`rowHeights`, `hiddenRows` and `hiddenColumns`** appear only for files that have them.
- **`hyperlinks`** are returned as stored, whatever their scheme. Check `url` before you render it as an `<a href>`.

## Limits for files you did not make

`new ExcelReader({ maxCells, maxPartBytes, maxTotalBytes, maxSheets })` sets limits for files you did not make. The size limits and `maxSheets` are checked before the sheets are inflated. `maxCells` is checked while a sheet is parsed, after its XML is inflated, so use `maxPartBytes` to bound memory. See [Handle errors and limits](../errors/).

## Rows as objects

To read a table as an array of objects instead of cells, see [Rows as objects](../objects/).
