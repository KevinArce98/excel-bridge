<div align="center">

<img src="https://raw.githubusercontent.com/KevinArce98/excel-bridge/main/assets/banner.svg" alt="excel-bridge — the lightweight Excel toolkit for TypeScript" width="100%" />

<br />

**Read and write `.xlsx` files in the browser and Node.js — with styling, formulas, dates and multi-sheet support, and no heavy dependencies.**

<br />

[![npm version](https://img.shields.io/npm/v/excel-bridge?logo=npm&label=npm&color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![downloads](https://img.shields.io/npm/dm/excel-bridge?label=downloads&color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![min+gzip](https://img.shields.io/bundlephobia/minzip/excel-bridge?label=min%2Bgzip&color=22c55e)](https://bundlephobia.com/package/excel-bridge)
[![CI](https://img.shields.io/github/actions/workflow/status/KevinArce98/excel-bridge/ci.yml?branch=main&label=CI&logo=github)](https://github.com/KevinArce98/excel-bridge/actions)
[![types](https://img.shields.io/npm/types/excel-bridge?color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![license](https://img.shields.io/npm/l/excel-bridge?color=22c55e)](./LICENSE)

<sub>[Quick Start](#quick-start) · [Why excel-bridge?](#why-excel-bridge) · [Guide](#guide) · [API Reference](#api-reference) · [Compatibility](#compatibility)</sub>

</div>

---

## Highlights

- **Read & write `.xlsx`** — parse existing workbooks or generate new ones from plain arrays.
- **Zero heavy dependencies** — no ExcelJS or SheetJS under the hood, just `fflate` + `fast-xml-parser`.
- **Tiny & tree-shakeable** — micro-package architecture ships only what you import (ESM + CJS).
- **TypeScript-first** — complete types and IntelliSense for every public API.
- **Cross-platform** — one API for the browser (`File`/`Blob`) and Node.js (`Buffer`).
- **Rich styling** — background, fonts, borders, alignment, wrapping and custom number formats.
- **Real spreadsheet features** — formulas, dates, merged cells, freeze panes, auto-width and data validation.
- **Multi-sheet workbooks** — build named sheets with per-sheet options in a single file.

## Installation

```bash
npm install excel-bridge
```

```bash
pnpm add excel-bridge
```

```bash
yarn add excel-bridge
```

## Quick Start

### Read a workbook

```typescript
import { ExcelBridge } from 'excel-bridge';

// Browser — from a file <input>
const file = document.querySelector<HTMLInputElement>('input[type="file"]')!.files![0];
const workbook = await ExcelBridge.readFromFile(file);

// Node.js — from a Buffer
import fs from 'node:fs';
const buffer = fs.readFileSync('data.xlsx');
const workbook = ExcelBridge.read(buffer);

// Every cell is typed. Dates come back as `Date`, formula cells expose `.formula`.
for (const row of workbook.sheets[0].data) {
  for (const cell of row) {
    console.log(cell.coordinate, cell.type, cell.value, cell.formula ?? '');
    // "B2" "date"   2024-01-15T00:00:00.000Z ""
    // "D2" "number" null                     "B2*C2"
  }
}
```

> Files produced by Excel or other libraries are read correctly: worksheets are resolved
> through their relationship ids, sparse rows keep their column alignment, and
> date-formatted cells are returned as `Date` objects.

### Write a workbook

```typescript
import { ExcelBridge } from 'excel-bridge';

const data = [
  ['Name', 'Age', 'City'],
  ['John', 25, 'New York'],
  ['Jane', 30, 'Los Angeles'],
];

// Browser — get a Blob to download
const blob = ExcelBridge.write(data);
const url = URL.createObjectURL(blob);

// Node.js — get a Buffer to write to disk
import fs from 'node:fs';
fs.writeFileSync('output.xlsx', ExcelBridge.writeBuffer(data));
```

## Why excel-bridge?

The `.xlsx` ecosystem is dominated by two large libraries. `excel-bridge` targets the common
case — **styled, multi-sheet reports read and written from typed data** — while staying small
enough to drop into a front-end bundle.

| | **excel-bridge** | ExcelJS | SheetJS (community `xlsx`) |
| --- | :---: | :---: | :---: |
| Read `.xlsx` | ✅ | ✅ | ✅ |
| Write `.xlsx` | ✅ | ✅ | ✅ |
| Cell styling (color, font, border) | ✅ | ✅ | ⚠️ Pro edition |
| Formulas | ✅ | ✅ | ✅ |
| Merged cells & freeze panes | ✅ | ✅ | ✅ |
| First-class TypeScript types | ✅ | ✅ | ✅ |
| ESM **and** CJS, tree-shakeable | ✅ | ⚠️ | ✅ |
| Heavy runtime dependencies | **None** | Several | None |
| Bundle footprint | **Tiny** ¹ | Large ¹ | Large ¹ |

<sub>¹ Comparisons refer to the free/community distributions. Check current, exact sizes on
Bundlephobia: [excel-bridge](https://bundlephobia.com/package/excel-bridge) ·
[exceljs](https://bundlephobia.com/package/exceljs) · [xlsx](https://bundlephobia.com/package/xlsx).</sub>

## Guide

- [Multi-sheet workbooks](#multi-sheet-workbooks)
- [Styling cells](#styling-cells)
- [Extended cell styles](#extended-cell-styles)
- [Formulas & dates](#formulas--dates)
- [Merged cells & layout](#merged-cells--layout)
- [Shared strings (opt-in)](#shared-strings-opt-in)
- [Coordinate helpers](#coordinate-helpers)

### Multi-sheet workbooks

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

const salesSheet = {
  data: [
    ['Product', 'Price', 'Quantity', 'Total'],
    ['Laptop', 999.99, 5, '=B2*C2'],
    ['Mouse', 29.99, 20, '=B3*C3'],
    ['Keyboard', 79.99, 10, '=B4*C4'],
  ],
  styles: {
    '0-0': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-1': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-2': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-3': { background: '#4472C4', bold: true, color: '#FFFFFF' },
  },
  options: { name: 'Sales Report', freezePane: { row: 1 }, autoWidth: true },
};

const datesSheet = {
  data: [
    ['Event', 'Date', 'Days Until Today'],
    ['Launch', new Date(2024, 6, 15), '=TODAY()-B2'],
    ['Meeting', new Date(2024, 8, 20), '=TODAY()-B3'],
    ['Deadline', new Date(2024, 11, 31), '=TODAY()-B4'],
  ],
  options: { name: 'Timeline', autoWidth: true },
};

const buffer = writer.createWorkbookBuffer([salesSheet, datesSheet]);

import fs from 'node:fs';
fs.writeFileSync('report.xlsx', buffer);
```

### Styling cells

Style keys use `"<row>-<col>"` (zero-based) coordinates, so you can drive them from data.

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

const sheet = {
  data: [
    ['Product', 'Price', 'Stock', 'Status'],
    ['Laptop', 999.99, 15, 'Available'],
    ['Mouse', 29.99, 5, 'Low Stock'],
    ['Keyboard', 79.99, 0, 'Out of Stock'],
  ],
  styles: {
    // Header row
    '0-0': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-1': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-2': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-3': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    // Conditional highlights
    '1-2': { background: '#E2EFDA', color: '#006100' }, // in stock
    '2-2': { background: '#FFC7CE', color: '#9C0006' }, // low stock
    '3-3': { background: '#FFE6E6', color: '#C00000' }, // out of stock
  },
  options: { name: 'Inventory', freezePane: { row: 1 }, autoWidth: true },
};

const buffer = writer.createWorkbookBuffer([sheet]);
```

### Extended cell styles

Beyond background, bold, color and borders, cells support fonts, alignment and number formats:

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const sheet = {
  data: [
    ['Invoice', 1250.5],
    ['Tax', 237.6],
  ],
  styles: {
    // Title: large, italic, centered, wrapped
    '0-0': { bold: true, italic: true, fontSize: 14, fontName: 'Arial', align: 'center', wrapText: true },
    // Amounts: custom currency number format
    '0-1': { numberFormat: '#,##0.00' },
    '1-1': { numberFormat: '#,##0.00', verticalAlign: 'middle' },
  },
};

const buffer = writer.createWorkbookBuffer([sheet]);
```

### Formulas & dates

`Date` objects are converted to Excel serials automatically, and any string starting with `=`
is written as a formula. Formulas are recalculated by Excel when the file is opened.

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const projectSheet = {
  data: [
    ['Task', 'Start Date', 'End Date', 'Duration', 'Status'],
    ['Design', new Date(2024, 0, 15), new Date(2024, 1, 20), '=C2-B2', 'Completed'],
    ['Development', new Date(2024, 1, 21), new Date(2024, 4, 30), '=C3-B3', 'In Progress'],
    ['Testing', new Date(2024, 5, 1), new Date(2024, 5, 15), '=C4-B4', 'Planned'],
    ['', '', '', '', ''],
    ['Tasks Completed', '', '', '=COUNTIF(E2:E4,"Completed")', ''],
  ],
  options: { name: 'Project Timeline', freezePane: { row: 1 }, autoWidth: true },
};

const buffer = writer.createWorkbookBuffer([projectSheet]);
```

### Merged cells & layout

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const reportSheet = {
  data: [
    ['Q1 2024 Sales Report', '', '', ''],
    ['Product', 'January', 'February', 'March'],
    ['Laptops', 45000, 52000, 48000],
    ['Accessories', 12000, 15000, 13500],
    ['TOTAL', '=SUM(B3:B4)', '=SUM(C3:C4)', '=SUM(D3:D4)'],
  ],
  styles: {
    '0-0': { background: '#5B9BD5', bold: true, color: '#FFFFFF' },
    '4-0': { background: '#70AD47', bold: true, color: '#FFFFFF' },
  },
  mergeCells: ['A1:D1'], // merge the title row
  options: { name: 'Quarterly Report', freezePane: { row: 2 }, autoWidth: true },
};

const buffer = writer.createWorkbookBuffer([reportSheet]);
```

### Shared strings (opt-in)

Strings are written inline by default (simple and reliable). For workbooks with many repeated
strings, enable a shared-strings table to reduce file size:

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ sharedStrings: true });
const buffer = writer.createWorkbookBuffer([{ data }]);
```

### Coordinate helpers

```typescript
import { coordinateToIndex, indexToCoordinate } from 'excel-bridge';

coordinateToIndex('A1');    // { row: 0, col: 0 }
indexToCoordinate(0, 0);    // "A1"
```

## API Reference

### Entry points

| Export | Description |
| --- | --- |
| `ExcelBridge.read(buffer)` | Parse an `.xlsx` from a `Buffer`/`Uint8Array`. |
| `ExcelBridge.readFromFile(file)` | Parse an `.xlsx` from a browser `File`. |
| `ExcelBridge.write(data)` | Create an `.xlsx` `Blob` from a 2D array. |
| `ExcelBridge.writeBuffer(data)` | Create an `.xlsx` `Buffer` from a 2D array. |

### Classes

| Class | Description |
| --- | --- |
| `ExcelReader` | Full-featured `.xlsx` parser. |
| `ExcelWriter` | Multi-sheet `.xlsx` writer with styling, formulas and layout. |
| `StyleManager` | Deduplicated style registry used internally by the writer. |

### Utilities

| Function | Description |
| --- | --- |
| `coordinateToIndex(coord)` | `"A1"` → `{ row, col }`. |
| `indexToCoordinate(row, col)` | `{ row, col }` → `"A1"`. |
| `dateToExcelSerial(date)` | `Date` → Excel serial number. |
| `excelSerialToDate(serial)` | Excel serial number → `Date`. |
| `isDate(value)` | Type guard for valid `Date` objects. |
| `calculateColumnWidths(data)` | Compute optimal column widths. |

### Types

```typescript
// Values accepted in a cell. Dates become Excel serials; strings starting with "=" are formulas.
type CellValue = string | number | boolean | Date | null | undefined;

interface ExcelData {
  data: CellValue[][];
  styles?: Record<string, CellStyle>;
  validations?: CellValidation[];
  mergeCells?: string[];
  options?: SheetOptions;
}

interface SheetOptions {
  name?: string;
  freezePane?: { row?: number; col?: number };
  autoWidth?: boolean;
}

interface CellStyle {
  background?: string;
  border?: boolean;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSize?: number;
  fontName?: string;
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  /** Custom Excel number-format code, e.g. "0.00" or "#,##0". */
  numberFormat?: string;
}

interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'date' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
  /** Present when the cell holds a formula (without the leading "="). */
  formula?: string;
}

interface ExcelWriterOptions {
  creator?: string;
  title?: string;
  subject?: string;
  /** Write strings to a shared-strings table instead of inline. Default: false. */
  sharedStrings?: boolean;
}
```

## Compatibility

| Environment | Support |
| --- | --- |
| Node.js | `^20.19.0`, `^22.13.0`, or `>=24` |
| Browsers | Modern browsers with ES2022, `File` and `Blob` APIs |
| Module formats | ESM (`import`) and CommonJS (`require`) |
| Excel | Excel 2016+ for full feature compatibility |

### Known limitations

- **Inline strings by default** — enable a shared-strings table with `new ExcelWriter({ sharedStrings: true })` for smaller files with lots of repeated text.
- **Formulas recalculate on open** — formula cells are written without a cached value; Excel computes them on load (`fullCalcOnLoad`).

## Contributing

Issues and pull requests are welcome. To work on the library locally:

```bash
pnpm install
pnpm run build          # bundle ESM + CJS + types
pnpm run test           # run the test suite (watch)
pnpm run test:run       # run once (CI / pre-publish)
pnpm run lint           # ESLint
pnpm run format:check   # Prettier
```

See the [CHANGELOG](./CHANGELOG.md) for release notes.

## License

[MIT](./LICENSE) © [Kevin Arias](https://github.com/KevinArce98)
