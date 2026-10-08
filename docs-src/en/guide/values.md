---
title: Write values and formulas
description: What a cell can hold, how formulas, literal text and errors are written, and what the library checks before it writes.
group: Guide
groupOrder: 2
order: 1
---

# Write values and formulas

A sheet is an array of rows, and a row is an array of cells. Most cells are plain values. A few need an object, because they say something a plain value cannot.

## Plain values

| You write | The cell holds |
| --- | --- |
| `'Laptop'` | Text. A string is always text, even if it starts with `=`. |
| `999.99` | A number. It must be finite. |
| `true` | A boolean. |
| `new Date(2024, 0, 15)` | A date, stored as an Excel serial number. |
| `null` or `undefined` | An empty cell. |

```ts title="plain-values.ts"
import { ExcelWriter } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Item', 'Price', 'In stock', 'Added', 'Note'],
      ['Laptop', 999.99, true, new Date(2024, 0, 15), null],
      ['=not a formula', 29.99, false, new Date(2024, 1, 2), 'plain text'],
    ],
  },
]);
```

Because strings are always text, exporting user data cannot create a formula by accident. A name such as `=HYPERLINK("https://example.com")` is written as the text it is.

## Formulas

Write a formula with an object. Leave out the leading `=`.

```ts title="formulas.ts"
import { ExcelWriter } from 'excel-bridge';

const bytes = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Product', 'Price', 'Quantity', 'Total'],
      ['Laptop', 999.99, 5, { formula: 'B2*C2', result: 4999.95 }],
      ['Mouse', 29.99, 20, { formula: 'B3*C3' }],
      ['', '', 'Sum', { formula: 'SUM(D2:D3)' }],
    ],
  },
]);
```

- **`{ formula }`** writes a formula with no stored value. The workbook asks Excel to recalculate every formula when it opens the file.
- **`{ formula, result }`** also stores the value the formula last had. `result` is a string, a number, a boolean, a `Date` or an `{ error }`.

The library never calculates a formula. A reader that does not calculate, such as a previewer or a script, shows `result` when you provide one and an empty cell when you do not.

> [!WARNING]
> Never pass untrusted input to `{ formula }`. A formula runs when the file is opened. User text belongs in a plain string.

## Literal text and errors

`{ text }` writes text, exactly like a plain string. Use it when the same code must also run on excel-bridge 1.6, where a plain string that starts with `=` is a formula. Before 1.6 an object cell is written as `[object Object]`.

`{ error }` writes an error value. The seven classic errors are `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!` and `#N/A`. Any other error text throws.

```ts title="text-and-errors.ts"
import type { CellValue } from 'excel-bridge';

const row: CellValue[] = [{ text: '=SUM(A1:A3)' }, { error: '#N/A' }, { formula: '1/0', result: { error: '#DIV/0!' } }];
```

## Dates

A `Date` is written as local wall-clock time, so `new Date(2024, 0, 15)` is 15 January in every time zone. Two helpers convert between dates and Excel serial numbers: `dateToExcelSerial(date)` and `excelSerialToDate(serial)`.

> [!LIMIT]
> A UTC-midnight date such as `new Date('2024-01-15')` falls on the previous day in negative offsets. A local time that does not exist, inside a daylight-saving gap, moves forward.

To show a time of day or another layout, give the cell a date `numberFormat`. See [Style cells](../styling/).

## What the writer checks

The writers throw an [`ExcelBridgeError`](../errors/) with code `INVALID_INPUT` for the inputs below.

| Input | Result |
| --- | --- |
| `NaN`, `Infinity` or an invalid `Date` | Throws and names the cell, such as `Cell B2 holds NaN, which a worksheet cannot store`. |
| A string over 32,767 characters | Throws. |
| `{ formula: '' }` | Throws. |
| An error value outside the seven | Throws. |
| Control characters other than tab, line feed and carriage return | Removed silently. |

## Strings and file size

Strings are written inline by default, which is simple and reliable. For a workbook with many repeated strings, enable a shared strings table to make the file smaller.

```ts title="shared-strings.ts"
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ sharedStrings: true });
const bytes = writer.createWorkbookBuffer([
  { data: [['Status'], ['Open'], ['Open'], ['Done'], ['Open']] },
]);
```

The [streaming writer](../streaming/) ignores this option and always writes inline strings.
