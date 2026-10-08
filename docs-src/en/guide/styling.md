---
title: Style cells
description: Colour, fonts, number formats, alignment and per-side borders for any cell, keyed by row and column.
group: Guide
groupOrder: 2
order: 2
---

# Style cells

Give any cell a fill, a font, an alignment, a number format or a border. Styles live next to the data, keyed by `"<row>-<col>"` with zero-based indexes, so you can build them from the data itself.

```ts title="header-row.ts"
import { ExcelWriter } from 'excel-bridge';
import type { ExcelData } from 'excel-bridge';

const header = { background: '#4472C4', bold: true, color: '#FFFFFF' };

const sheet: ExcelData = {
  data: [
    ['Product', 'Price', 'Stock', 'Status'],
    ['Laptop', 999.99, 15, 'Available'],
    ['Mouse', 29.99, 5, 'Low stock'],
    ['Keyboard', 79.99, 0, 'Out of stock'],
  ],
  styles: {
    '0-0': header,
    '0-1': header,
    '0-2': header,
    '0-3': header,
    '1-2': { background: '#E2EFDA', color: '#006100' },
    '2-2': { background: '#FFC7CE', color: '#9C0006' },
    '3-3': { background: '#FFE6E6', color: '#C00000' },
  },
  options: { name: 'Inventory', freezePane: { row: 1 }, autoWidth: true },
};

const bytes = new ExcelWriter().createWorkbookBuffer([sheet]);
```

> [!TIP]
> Annotate the sheet as `ExcelData`. Without it, TypeScript widens literals such as `align: 'center'` to `string` and rejects the object.

## What a style can hold

| Property | Value | Effect |
| --- | --- | --- |
| `background` | colour | Solid fill. |
| `color` | colour | Font colour. |
| `bold`, `italic`, `underline` | `boolean` | Font weight, slant, underline. |
| `fontSize` | `number` | Points. Default 11. |
| `fontName` | `string` | Default `Calibri`. |
| `align` | `'left' \| 'center' \| 'right'` | Horizontal alignment. |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | Vertical alignment. |
| `wrapText` | `boolean` | Wrap long text inside the cell. |
| `numberFormat` | format code | For example `'#,##0.00'` or `'0%'`. |
| `border` | `boolean`, a line style or per-side entries | See [Borders](#borders). |

Colours are hex: `#RGB`, `#RRGGBB` or `#AARRGGBB`, with or without the `#`. Anything else, such as `red`, throws instead of writing an invalid value.

```ts title="invoice.ts"
import type { ExcelData } from 'excel-bridge';

const invoice: ExcelData = {
  data: [
    ['Invoice', 1250.5],
    ['Tax', 237.6],
  ],
  styles: {
    '0-0': { bold: true, italic: true, fontSize: 14, fontName: 'Arial', align: 'center', wrapText: true },
    '0-1': { numberFormat: '#,##0.00' },
    '1-1': { numberFormat: '#,##0.00', verticalAlign: 'middle' },
  },
};
```

## Dates take styles too

A `Date` cell uses the built-in short date format. Give its style a date `numberFormat` to show the time of day or another layout. A `numberFormat` that is not a date format is ignored on a `Date` cell.

```ts title="dates.ts"
import type { ExcelData } from 'excel-bridge';

const events: ExcelData = {
  data: [['Launch', new Date(2024, 0, 15, 13, 45)]],
  styles: { '0-1': { bold: true, numberFormat: 'yyyy-mm-dd hh:mm' } },
};
```

> [!LIMIT]
> A `Date` is written as local wall-clock time: `new Date(2024, 0, 15)` is 15 January in every time zone. A UTC-midnight date such as `new Date('2024-01-15')` falls on the previous day in negative offsets, and a local time that does not exist (inside a daylight-saving gap) moves forward.

## Borders

`border: true` draws a thin black box on all four sides. For anything else, pass a line style, a `{ style, color? }` pair, or an object with one entry per side.

```ts title="borders.ts"
import type { ExcelData } from 'excel-bridge';

const report: ExcelData = {
  data: [
    ['Region', 'Revenue'],
    ['North', 1200],
    ['South', 480],
  ],
  styles: {
    '0-0': { bold: true, border: { bottom: 'medium' } },
    '0-1': { bold: true, border: { bottom: 'medium' } },
    '1-1': { border: { left: { style: 'thin', color: '#CCCCCC' } } },
    '2-0': { border: 'thin' },
    '2-1': { border: { top: 'thin', bottom: { style: 'double', color: '#CC0000' } } },
  },
};
```

- **Accepted values:** `true` or `false`; a line style for all four sides; `{ style, color? }` for all four sides; or `{ left?, right?, top?, bottom? }`, each a line style or `{ style, color? }`.
- **Line styles (13):** `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`.
- **Colours:** `#RGB`, `#RRGGBB` or `#AARRGGBB`. Without a colour the line is black.
- **Errors:** an unknown line style or a colour such as `red` throws and names the value.
- **Merged cells:** the file stores borders per cell, and the writer does not copy a border across a merged range. Set the side on each edge cell and pad the data with `null` so those cells exist. A style on a cell past the end of its row is ignored.

> [!LIMIT]
> No diagonal borders, and no borders inside conditional formats.

## Conditional formatting without a style per cell

Highlight cells by value or formula, or with a colour scale, using [rules](../rules/) instead of one style per cell.
