<div align="center">

<img src="https://raw.githubusercontent.com/KevinArce98/excel-bridge/main/assets/banner.svg" alt="excel-bridge — the lightweight Excel toolkit for TypeScript" width="100%" />

<br />

**Read and write styled `.xlsx` files in the browser and Node.js, without shipping ExcelJS or SheetJS.**

Import a class and your bundler ships only that part: the writer adds 12.5 KB min+gzip.

<br />

[![Live demo](https://img.shields.io/badge/demo-live-22c55e?labelColor=1e293b)](https://kevinarce98.github.io/excel-bridge/)
[![npm version](https://img.shields.io/npm/v/excel-bridge?logo=npm&label=npm&color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![ExcelWriter size](https://img.shields.io/badge/ExcelWriter-12.5%20KB%20min%2Bgzip-22c55e?labelColor=1e293b)](#bundle-size)
[![CI](https://img.shields.io/github/actions/workflow/status/KevinArce98/excel-bridge/ci.yml?branch=main&label=CI&logo=github)](https://github.com/KevinArce98/excel-bridge/actions)
[![types](https://img.shields.io/npm/types/excel-bridge?color=22c55e)](https://www.npmjs.com/package/excel-bridge)
[![license](https://img.shields.io/npm/l/excel-bridge?color=22c55e)](./LICENSE)

<sub>[Live Demo](https://kevinarce98.github.io/excel-bridge/) · [Quick Start](#quick-start) · [Why excel-bridge?](#why-excel-bridge) · [Bundle size](#bundle-size) · [Guide](#guide) · [API Reference](#api-reference) · [Compatibility](#compatibility)</sub>

</div>

---

## Highlights

- **Small and tree-shakeable.** `ExcelWriter` adds 12.5 KB min+gzip. Import only the class you need ([sizes](#bundle-size)). ESM and CJS builds.
- **Read and write.** Cell styles (fill, font, one box border, alignment, number formats), formulas, dates, merged cells, freeze panes, column widths, **conditional formatting**, data validation, **autofilters**, **hyperlinks**, hidden sheets and multi-sheet workbooks.
- **One API for the browser and Node.js.** `Blob` and `File` in the browser, `Uint8Array` in Node.js (a `Buffer` is one). Reading and writing buffers is synchronous; reading a `File` and streaming are async.
- **TypeScript-first.** Complete types and IntelliSense for every public export.
- **Built for large exports.** A streaming writer produces million-row files from sync or async iterables, so the rows never all live in memory at once.
- **Edit what you wrote.** `Workbook` loads, edits and saves files the library can model ([what a round trip keeps](#what-a-round-trip-keeps)).
- **Two direct dependencies.** `fflate` for zip and `fast-xml-parser` for reading. The writer bundles no parser code.
- **Signed releases.** Every version is published to npm with [provenance](https://docs.npmjs.com/generating-provenance-statements).

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

> **[Try the live demo →](https://kevinarce98.github.io/excel-bridge/)** Build a styled workbook and download it, or drop in an `.xlsx` and see it parsed, all in your browser. Prefer Node.js? Run it on [RunKit](https://npm.runkit.com/excel-bridge).

## Quick Start

Import the class you need: `ExcelWriter`, `ExcelReader` or [`Workbook`](#high-level-workbook-api).
Your bundler ships only that part of the library ([sizes](#bundle-size)).

### Write a workbook

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const header = { bold: true, background: '#4472C4', color: '#FFFFFF' };

const sheet = {
  data: [
    ['Name', 'Age', 'City'],
    ['John', 25, 'New York'],
    ['Jane', 30, 'Los Angeles'],
  ],
  styles: { '0-0': header, '0-1': header, '0-2': header },
  options: { name: 'People', freezePane: { row: 1 }, autoWidth: true },
};

// Browser: get a Blob to download
const blob = writer.createWorkbook([sheet]);
const url = URL.createObjectURL(blob);

// Node.js: get a Uint8Array to write to disk
import fs from 'node:fs';
fs.writeFileSync('output.xlsx', writer.createWorkbookBuffer([sheet]));
```

`createWorkbookBuffer` returns a plain `Uint8Array`, not a `Buffer`. Frameworks that look for a
`Buffer` need a wrapper, for example `res.send(Buffer.from(bytes))` in Express 4.

### Read a workbook

```typescript
import { ExcelReader } from 'excel-bridge';

const reader = new ExcelReader();

// Node.js: from a Buffer or Uint8Array (synchronous)
import fs from 'node:fs';
const workbook = reader.parseFromBuffer(fs.readFileSync('data.xlsx'));

// Browser: from a File, such as the one an <input type="file"> gives you (async)
const uploaded = await reader.parseFromFile(file);

// Every cell carries a type tag. Dates come back as `Date`, formula cells expose `.formula`.
for (const row of workbook.sheets[0].data) {
  for (const cell of row) {
    console.log(cell.coordinate, cell.type, cell.value, cell.formula ?? '');
    // "B2" "date"   2024-01-15T00:00:00.000Z ""   (printed in UTC; a Date holds local wall-clock time)
    // "D2" "empty"  null                     "B2*C2"   (a formula without a cached value)
  }
}
```

The reader finds worksheets through the file's relationship ids, so files from Excel and other
libraries open wherever they store their sheets. Date-formatted cells come back as `Date` objects,
and error cells come back as `type: 'error'` with their text, for example `#DIV/0!`. `sheet.data`
lists the rows present in the file in order; use `cell.rowIndex` for the position. A few cases are
not read correctly yet, such as the 1904 date system; see [Known limitations](#known-limitations).

### Convenience entry point

`ExcelBridge` puts the reader, the writer and the helpers on a single object, which suits Node.js
scripts and quick prototypes:

```typescript
import { ExcelBridge } from 'excel-bridge';

const workbook = ExcelBridge.read(buffer);
const blob = ExcelBridge.write([
  ['Name', 'Age'],
  ['John', 25],
]);
```

It also offers `ExcelBridge.readFromFile(file)` for browser `File`s and `ExcelBridge.writeBuffer(data)`
for Node.js ([all entry points](#entry-points)).

> **Bundle size:** `ExcelBridge` is 40.1 KB min+gzip, against 12.5 KB for `ExcelWriter`, because bundlers keep an object whole: even a lone `ExcelBridge.write` call ships the reader too. In browser code, prefer the named imports.

## Why excel-bridge?

The `.xlsx` ecosystem is dominated by two large libraries. `excel-bridge` targets the common case,
**styled, multi-sheet reports read and written from typed data**, and stays small enough to drop
into a front-end bundle.

| | **excel-bridge** | ExcelJS | SheetJS (community `xlsx`) |
| --- | :---: | :---: | :---: |
| Read `.xlsx` | ✅ | ✅ | ✅ |
| Write `.xlsx` | ✅ | ✅ | ✅ |
| Cell styling (color, font, one box border) | ✅ | ✅ | ⚠️ Pro edition |
| Conditional formatting | ✅ | ✅ | ⚠️ Pro edition |
| Formulas | ✅ | ✅ | ✅ |
| Merged cells | ✅ | ✅ | ✅ |
| Freeze panes | ✅ | ✅ | ❌ not in `xlsx` 0.18.5 |
| Streaming writer | ✅ | ✅ | ⚠️ Pro edition |
| First-class TypeScript types | ✅ | ✅ | ✅ |
| ESM **and** CJS, tree-shakeable | ✅ | ⚠️ CJS-first | ✅ |
| Direct runtime dependencies ² | 2 | 9 | 7 |
| Bundle size to write a file ¹ | **12.5 KB** | 272.1 KB | 95.8 KB |

<sub>¹ Minified + gzipped code that a browser bundle needs to write an `.xlsx`: `ExcelWriter`, ExcelJS's default browser build (not tree-shakeable) and SheetJS `utils` + `write` from npm `xlsx@0.18.5`, bundled with esbuild. See [Bundle size](#bundle-size) for the method, reader sizes and more libraries. Bundlephobia measures each whole package with its own toolchain, so its numbers differ: [excel-bridge](https://bundlephobia.com/package/excel-bridge) · [exceljs](https://bundlephobia.com/package/exceljs) · [xlsx](https://bundlephobia.com/package/xlsx).</sub>

<sub>² Declared in each `package.json` (`exceljs@4.4.0`, `xlsx@0.18.5` on npm), checked 2026-10-07. `xlsx@0.18.5` is the last version published to npm and has known security advisories.</sub>

### Bundle size

Bundlers keep the exports you import and drop the rest. This is what each entry point adds to a
browser bundle:

| Import from `excel-bridge` | min+gzip |
| --- | ---: |
| `createExcelWorkbookStream` | 11.9 KB |
| `ExcelWriter` | 12.5 KB |
| `ExcelReader` | 28.2 KB |
| `Workbook` (reader + writer) | 39.9 KB |
| `ExcelBridge` (convenience object) | 40.1 KB |
| Everything | 42.6 KB |

The same measurement for other libraries (2026-10-07):

| Library | Import | min+gzip |
| --- | --- | ---: |
| hucre 1.2.0 | `writeXlsx` | 41.6 KB |
| hucre 1.2.0 | `readXlsx` | 41.1 KB |
| hucre 1.2.0 | `XlsxStreamWriter` (from `hucre/xlsx`) | 12.1 KB |
| hucre 1.2.0 | `streamXlsxRows` (from `hucre/xlsx`) | 17.4 KB |
| @mitresthen/excelents 1.0.1 | whole main entry, reader and writer | 12.2 KB |
| read-excel-file 9.3.10 | `read-excel-file/browser`, read only | 16.7 KB |
| write-excel-file 4.1.1 | `write-excel-file/universal`, write only | 19.8 KB |
| SheetJS (`xlsx` 0.18.5 on npm) | `utils` + `write` | 95.8 KB |
| ExcelJS 4.4.0 | default browser build, not tree-shakeable | 272.1 KB |

**`excel-bridge` is not the smallest on every row.** `ExcelWriter` writes styles, conditional
formatting, data validation, autofilters and hyperlinks in 12.5 KB. Some libraries are smaller for
one job, or give up a feature to get there (the `excelents` entry has no conditional formatting at
1.0.1). Compare what you get for the bytes, not the bytes alone.

Tree-shaking uses the ESM build, which bundlers pick for `import`. `require('excel-bridge')` loads
the whole CommonJS build.

<details>
<summary>How these numbers were measured</summary>

- **Toolchain:** esbuild 0.28.2 (`--bundle --minify --platform=browser --format=esm`), measured 2026-10-07.
- **Entry:** each row bundles a one-line `export { … } from '<package>'` file, then gzips the output with Node's zlib at the default level. 1 KB = 1,000 bytes.
- **Sources:** the excel-bridge rows come from this repository's build with the dependency versions in `pnpm-lock.yaml` (a fresh install that resolves newer `fast-xml-parser` patch releases can add about 50 bytes). The other libraries were installed in a scratch directory.
- **Margin:** gzip implementations differ by about 1%. macOS `gzip` comes out slightly smaller.
- **Not measured:** newer SheetJS Community Edition builds, which ship from the SheetJS CDN.
- **Reproduce:** run [`pnpm run size`](./benchmarks/README.md#bundle-size). CI runs `pnpm run size:check`, which fails when an excel-bridge row drifts more than 100 bytes from the measured size.

</details>

### Performance

Writing **50,000 rows × 10 columns** (median of 5 runs, Node 24.19, Apple M4, measured 2026-10-07):

| Library | Write time | Output size |
| --- | ---: | ---: |
| **excel-bridge** | **560 ms** | **2.41 MiB** |
| excel-bridge (streaming) | 590 ms | 2.48 MiB |
| hucre | 390 ms | 2.79 MiB |
| xlsx / SheetJS (default, no compression) | 510 ms | 18.23 MiB |
| xlsx / SheetJS (`compression: true`) | 570 ms | 6.26 MiB |
| exceljs | 1490 ms | 2.82 MiB |

- **Against ExcelJS:** about 2.7× faster.
- **Against hucre:** hucre writes this workload faster, in about 390 ms.
- **Against SheetJS:** similar time, 510 ms with its default uncompressed output and 570 ms with `compression: true`. Its file is about 7.5× larger by default and about 2.6× larger with compression.
- **Reading is slower:** `ExcelReader` takes 1.4 to 1.9 s for 50,000 × 10 cells and holds the whole file in memory (about 2 GB peak for 200,000 × 10; Node 24.19, Apple M4, files written by ExcelJS).

<sub>Times vary by about 10% between runs and by machine. Reproduce them with [`pnpm run bench`](./benchmarks/README.md#write-speed).</sub>

### When another library fits better

`excel-bridge` writes reports and reads them back. It does not cover everything:

- **Images, charts, comments, Excel tables, pivot tables, print setup, rich text, per-side borders, row heights, hidden rows and columns.** None of these are supported. ExcelJS and hucre cover many of them.
- **Editing a file and keeping everything you did not touch** (templates with charts or macros). `Workbook` rebuilds the file from what it models, so anything else is dropped. hucre's `openXlsx`/`saveXlsx` keeps the parts it does not model.
- **Reading very large files without loading them whole.** `ExcelReader` has no streaming mode. hucre has `streamXlsxRows`.
- **CSV, ODS or legacy `.xls`/`.xlsb`.** Use hucre or SheetJS.
- **Schema-validated imports** that map rows to typed objects. Use read-excel-file.

## Guide

- [High-level `Workbook` API](#high-level-workbook-api)
- [Multi-sheet workbooks](#multi-sheet-workbooks)
- [Styling cells](#styling-cells)
- [Extended cell styles](#extended-cell-styles)
- [Formulas & dates](#formulas--dates)
- [Merged cells & layout](#merged-cells--layout)
- [Conditional formatting](#conditional-formatting)
- [Data validation](#data-validation)
- [AutoFilter](#autofilter)
- [Hyperlinks](#hyperlinks)
- [Streaming large workbooks](#streaming-large-workbooks)
- [Shared strings (opt-in)](#shared-strings-opt-in)
- [Reading in depth](#reading-in-depth)
- [Coordinate helpers](#coordinate-helpers)

### High-level `Workbook` API

`Workbook` builds a workbook, or loads one, edits it and saves it back, without rebuilding sheet
data by hand. It rebuilds the file from what the library models, so it suits files this library
wrote and simple files from other tools. It does not keep images, charts, comments, Excel tables,
pivot tables, defined names, print setup, themes, macros, row heights or hidden rows and columns of
a file you load.

```typescript
import fs from 'node:fs';
import { Workbook } from 'excel-bridge';

// Start from scratch…
const wb = Workbook.create();
wb.addSheet('Sales', [
  ['Product', 'Price', 'Qty'],
  ['Laptop', 999.99, 5],
]);
wb.setCellStyle('Sales', 0, 0, { bold: true, background: '#4472C4', color: '#FFFFFF' });
wb.setFreezePane('Sales', { row: 1 });
wb.setMetadata({ creator: 'My App', title: 'Sales Report' });
fs.writeFileSync('sales.xlsx', wb.toBuffer());

// …or load, edit and save an existing workbook
const existing = Workbook.fromBuffer(fs.readFileSync('sales.xlsx'));
existing.setCellValue('Sales', 1, 1, 899.99); // drop the laptop price
const blob = existing.toBlob(); // in the browser
```

Load with `Workbook.fromBuffer(buffer)`, or `await Workbook.fromFile(file)` in the browser. Methods
on an instance:

| Area | Methods |
| --- | --- |
| Sheets | `addSheet`, `renameSheet`, `removeSheet`, `getSheetNames`, `getSheetData`, `getSheetState`, `setSheetState` |
| Cells | `getCellValue`, `setCellValue`, `getCellStyle`, `setCellStyle` |
| Layout | `setMergeCells`, `setFreezePane`, `setColumnWidths`, `setAutoWidth` |
| Rules and links | `addValidation`, `addConditionalFormat`, `setAutoFilter`, `getAutoFilter`, `removeAutoFilter`, `setHyperlink`, `getHyperlinks`, `removeHyperlink` |
| Document | `getMetadata`, `setMetadata` |
| Output | `toBuffer`, `toBlob` |

#### What a round trip keeps

`Workbook.fromBuffer` and `fromFile` restore:

- data at its row and column, styles, merges, freeze panes and column widths
- **data validations, conditional formatting rules, autofilter ranges, hyperlinks and sheet visibility**

They change or drop:

- Filter criteria and sort state set in Excel.
- Links the writer does not accept (anything but `http:`, `https:`, `mailto:` or a location in the workbook), and validations of a type the writer does not know. `ExcelReader` still returns the links.
- Validation input and error messages, the error style (stop, warning, information) and the "show message" switches. Every saved rule shows both messages and blocks bad entries.
- Error cells, which are saved as text.
- Date cells, which lose their style and number format.
- Colours other than RGB, and built-in number formats such as `0%`. Only RGB colours and custom number formats are read.
- Border detail: every border becomes a thin black box on all four sides.
- Shared-formula followers and text cells that start with `=`, which are saved as formulas.
- Rich text, which is flattened, and split panes, which become freeze panes.
- Conditional formatting rules other than `cellIs`, `expression` and `colorScale`.
- Dates in the 1904 date system, which are read incorrectly.

After loading a file with blank rows, `getSheetData(name)` has holes at those rows: `length` is the
last row plus one, `for...of` yields `undefined` for a hole and `JSON.stringify` writes `null`. A
loaded sheet whose name Excel would reject (for example more than 31 characters) loads, but saving
throws until you give it a valid name with `renameSheet`.

### Multi-sheet workbooks

```typescript
import fs from 'node:fs';
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

const header = { background: '#4472C4', bold: true, color: '#FFFFFF' };

const salesSheet = {
  data: [
    ['Product', 'Price', 'Quantity', 'Total'],
    ['Laptop', 999.99, 5, '=B2*C2'],
    ['Mouse', 29.99, 20, '=B3*C3'],
    ['Keyboard', 79.99, 10, '=B4*C4'],
  ],
  styles: { '0-0': header, '0-1': header, '0-2': header, '0-3': header },
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
fs.writeFileSync('report.xlsx', buffer);
```

Sheet names must be 1 to 31 characters, with none of `\ / ? * [ ] :` and no control characters, no
apostrophe at either end, and unique when case is ignored. The writers, `Workbook.addSheet` and
`Workbook.renameSheet` throw otherwise, and so does writing a workbook without sheets.

### Styling cells

Style keys use `"<row>-<col>"` (zero-based) coordinates, so you can drive them from data.

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

const header = { background: '#4472C4', bold: true, color: '#FFFFFF' };

const sheet = {
  data: [
    ['Product', 'Price', 'Stock', 'Status'],
    ['Laptop', 999.99, 15, 'Available'],
    ['Mouse', 29.99, 5, 'Low Stock'],
    ['Keyboard', 79.99, 0, 'Out of Stock'],
  ],
  styles: {
    // Header row
    '0-0': header,
    '0-1': header,
    '0-2': header,
    '0-3': header,
    // Highlights
    '1-2': { background: '#E2EFDA', color: '#006100' }, // in stock
    '2-2': { background: '#FFC7CE', color: '#9C0006' }, // low stock
    '3-3': { background: '#FFE6E6', color: '#C00000' }, // out of stock
  },
  options: { name: 'Inventory', freezePane: { row: 1 }, autoWidth: true },
};

const buffer = writer.createWorkbookBuffer([sheet]);
```

Colours are hex: `#RGB`, `#RRGGBB` or `#AARRGGBB`, with or without the `#`. Anything else, such as
`red`, throws instead of being written as an invalid value like `FFRREEDD`.

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
    // Amounts: custom number format
    '0-1': { numberFormat: '#,##0.00' },
    '1-1': { numberFormat: '#,##0.00', verticalAlign: 'middle' },
  },
};

const buffer = writer.createWorkbookBuffer([sheet]);
```

### Formulas & dates

`Date` objects are converted to Excel serials automatically, and any string starting with `=` is
written as a formula. Excel recalculates formulas when it opens the file.

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

Numbers must be finite. `NaN`, `Infinity` and invalid `Date` values throw an error that names the
cell, such as `Cell B2 holds NaN, which a worksheet cannot store`, so a failed calculation never
ends up as a bad cell in an export.

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
  // Fixed widths instead of autoWidth
  options: { name: 'Quarterly Report', freezePane: { row: 2 }, columnWidths: [24, 12, 12, 12] },
};

const buffer = writer.createWorkbookBuffer([reportSheet]);
```

### Conditional formatting

Attach `conditionalFormats` to a sheet to highlight cells by value, by formula, or with a color
scale.

```typescript
import { ExcelWriter } from 'excel-bridge';
import type { ConditionalFormat } from 'excel-bridge';

const conditionalFormats: ConditionalFormat[] = [
  // Highlight low stock in red
  {
    type: 'cellValue',
    range: 'C2:C4',
    operator: 'lessThan',
    value: 10,
    style: { background: '#FFC7CE', color: '#9C0006' },
  },
  // Flag rows where a formula is true
  {
    type: 'expression',
    range: 'A2:D4',
    formula: '$D2="Out of Stock"',
    style: { background: '#FFE6E6', bold: true },
  },
  // Three-color scale across a numeric range
  {
    type: 'colorScale',
    range: 'B2:B4',
    colors: ['#F8696B', '#FFEB84', '#63BE7B'],
  },
];

const writer = new ExcelWriter();
const buffer = writer.createWorkbookBuffer([
  {
    data: [
      ['Product', 'Price', 'Stock', 'Status'],
      ['Laptop', 999.99, 15, 'Available'],
      ['Mouse', 29.99, 5, 'Low Stock'],
      ['Keyboard', 79.99, 0, 'Out of Stock'],
    ],
    conditionalFormats,
  },
]);
```

### Data validation

Use the typed `dataValidation` builders instead of hand-writing raw rule strings. Each returns a
`CellValidation` you can drop into a sheet's `validations` array (or `Workbook.addValidation`).

```typescript
import { ExcelWriter, dataValidation } from 'excel-bridge';

const writer = new ExcelWriter();
const buffer = writer.createWorkbookBuffer([
  {
    data: [['Status', 'Priority', 'Score', 'Due']],
    validations: [
      dataValidation.list('A2:A100', ['Open', 'In Progress', 'Done']), // dropdown
      dataValidation.wholeNumber('B2:B100', 'between', 1, 5),
      dataValidation.decimal('C2:C100', 'greaterThanOrEqual', 0),
      dataValidation.dateBetween('D2:D100', new Date(2024, 0, 1), new Date(2024, 11, 31)),
    ],
  },
]);
```

Builders: `list(range, values)`, `wholeNumber`, `decimal`, `textLength` (each
`(range, operator, value, value2?)`) and `dateBetween(range, start, end)`. Operators are
`between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`,
`lessThanOrEqual`.

### AutoFilter

Add filter dropdowns to a header row with `options.autoFilter`. Cover the data rows too, the way
Excel saves a filter:

```typescript
import { ExcelWriter } from 'excel-bridge';

const rows = [
  ['Region', 'Rep', 'Revenue'],
  ['North', 'Ann', 1200],
  ['South', 'Bob', 480],
];

const buffer = new ExcelWriter().createWorkbookBuffer([
  {
    data: rows,
    options: { freezePane: { row: 1 }, autoFilter: { range: `A1:C${rows.length}` } },
  },
]);
```

A sheet holds one filter. The writer also adds the hidden `_xlnm._FilterDatabase` name that Excel
writes and LibreOffice reads the range from. Filter criteria and sort state are not written or
read, so the file opens with every row visible. On a `Workbook`, use `setAutoFilter`,
`getAutoFilter` and `removeAutoFilter`.

### Hyperlinks

Attach links through a sheet's `hyperlinks` array. The cell keeps the text from `data`; the link
sets where a click goes. The `hyperlink` builders cover the three kinds:

```typescript
import { ExcelWriter, hyperlink } from 'excel-bridge';

const buffer = new ExcelWriter().createWorkbookBuffer([
  {
    data: [
      ['Resource', 'Contact', 'Details'],
      ['Docs', 'Email the team', 'See Q1'],
    ],
    hyperlinks: [
      hyperlink.url('A2', 'https://example.com/docs', { tooltip: 'Open the docs' }),
      hyperlink.email('B2', 'team@example.com', { subject: 'Report question' }),
      hyperlink.internal('C2', 'Q1 Sales', 'A1'),
    ],
    options: { name: 'Links' },
  },
  { data: [['Q1']], options: { name: 'Q1 Sales' } },
]);
```

A link is also a plain object: `{ range: 'A2', url: 'https://…' }` or
`{ range: 'C2', location: "'Q1 Sales'!A1" }`, with optional `tooltip` and `display`. A leading `#`
in `location` is dropped.

- **Look:** the first cell of each link gets Excel's hyperlink color (`#0563C1`) and an underline.
  A `color` or `underline` set in that cell's own style wins, each on its own.
- **Allowed URLs:** `http:`, `https:` and `mailto:`, with spaces and quotes percent-encoded.
  Anything else throws: exports often carry user data, and `file:` links or custom protocol
  handlers are not safe to click.
- **Limits:** one link per range, 65,530 links per sheet and 2,079 characters per address. Excel
  shows the first 255 characters of a tooltip.
- **Workbook:** `setHyperlink` replaces any link on the same range, `getHyperlinks` returns copies
  and `removeHyperlink` also removes the link look.

### Streaming large workbooks

For exports too large to hold in memory, `createExcelWorkbookStream` yields the `.xlsx` as
`Uint8Array` chunks. Rows come from a **sync or async iterable**, so they never all live in memory
at once.

```typescript
import { createExcelWorkbookStream, streamToBuffer } from 'excel-bridge';
import fs from 'node:fs';

// Rows are produced lazily. Here, a million of them.
function* generateRows() {
  yield ['Id', 'Name', 'Value'];
  for (let i = 1; i <= 1_000_000; i++) {
    yield [i, `Row ${i}`, i * 2];
  }
}

// Option A: pipe chunks straight to disk, nothing is buffered whole
const out = fs.createWriteStream('big.xlsx');
for await (const chunk of createExcelWorkbookStream([{ name: 'Data', rows: generateRows() }])) {
  out.write(chunk);
}
out.end();

// Option B: collect into a single Uint8Array when you still need a buffer
const buffer = await streamToBuffer(
  createExcelWorkbookStream([{ name: 'Data', rows: generateRows() }])
);
```

A streaming sheet (`StreamingSheetInput`) supports `name`, `rows`, `styles`, `freezePane`,
`columnWidths`, `mergeCells`, `autoFilter` and `hyperlinks`. Pass the final filter range and links
up front: they are validated, and the range is written to the workbook part, before the first row.
It does **not** support `autoWidth`, `validations` or `conditionalFormats`. Use
`ExcelWriter` or `Workbook` when you need those.

### Shared strings (opt-in)

Strings are written inline by default, which is simple and reliable. For workbooks with many
repeated strings, enable a shared-strings table to reduce file size:

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ sharedStrings: true });
const buffer = writer.createWorkbookBuffer([
  { data: [['Status'], ['Open'], ['Open'], ['Done'], ['Open']] },
]);
```

### Reading in depth

`ExcelReader` returns the workbook, not just cell values. Each `ParsedSheet` also exposes its
layout, and the workbook carries document metadata.

```typescript
import { ExcelReader } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(buffer);

const sheet = workbook.sheets[0];
sheet.name; // "Sales Report"
sheet.state; // "hidden" | "veryHidden" when the sheet is hidden, otherwise undefined
sheet.data; // the rows present in the file, in order; use cell.rowIndex for the position
sheet.styles; // Record<"row-col", CellStyle>
sheet.mergeCells; // ["A1:D1", ...]
sheet.freezePane; // { row?: number; col?: number }
sheet.columnWidths; // number[]
sheet.validations; // CellValidation[]: type, operator, formulas and allowBlank as stored
sheet.conditionalFormats; // ConditionalFormat[]: rules the writer produces, read back

workbook.metadata; // { created?, modified?, creator?, title?, subject? }
```

Sheets with a filter or links also carry `sheet.autoFilter` (`{ range }`) and `sheet.hyperlinks`
(`Hyperlink[]`). The reader returns every link as stored, whatever its scheme, so check `url`
before rendering it as an `<a href>`.

### Coordinate helpers

```typescript
import { coordinateToIndex, indexToCoordinate } from 'excel-bridge';

coordinateToIndex('A1');    // { row: 0, col: 0 }
indexToCoordinate(0, 0);    // "A1"
```

## API Reference

### Entry points

`ExcelBridge` groups the most common calls on one object. It is the simplest way to start, but it
bundles as a unit. For the smallest browser bundles, import the [classes](#classes) and
[functions](#functions) below directly ([sizes](#bundle-size)).

| Export | Description |
| --- | --- |
| `ExcelBridge.read(buffer)` | Parse an `.xlsx` from a `Buffer`/`Uint8Array` (synchronous). |
| `ExcelBridge.readFromFile(file)` | Parse an `.xlsx` from a browser `File` (async). |
| `ExcelBridge.write(data)` | Create an `.xlsx` `Blob` from a 2D array. |
| `ExcelBridge.writeBuffer(data)` | Create an `.xlsx` `Uint8Array` from a 2D array. |
| `ExcelBridge.Reader` / `.Writer` / `.Workbook` | The `ExcelReader`, `ExcelWriter` and `Workbook` classes. |
| `ExcelBridge.coordinateToIndex` / `.indexToCoordinate` | The coordinate helpers below. |

### Classes

| Class | Description |
| --- | --- |
| `Workbook` | High-level load / edit / save API over the reader and writer. |
| `ExcelReader` | `.xlsx` parser for the features the writer produces and the common ones from other tools. |
| `ExcelWriter` | Multi-sheet `.xlsx` writer with styling, formulas and layout. |
| `StyleManager` | Deduplicated style registry used internally by the writer. |

### Functions

| Function | Description |
| --- | --- |
| `createExcelWorkbookStream(sheets, options?)` | Async generator yielding `.xlsx` chunks for large exports. |
| `streamToBuffer(stream)` | Collect a workbook stream into a single `Uint8Array`. |
| `dataValidation.*` | Typed builders (`list`, `wholeNumber`, `decimal`, `textLength`, `dateBetween`) returning `CellValidation`. |
| `hyperlink.*` | Builders (`url`, `email`, `internal`) returning a `Hyperlink`. |
| `coordinateToIndex(coord)` | `"A1"` → `{ row, col }`. |
| `indexToCoordinate(row, col)` | `{ row, col }` → `"A1"`. |
| `dateToExcelSerial(date)` | `Date` → Excel serial number. |
| `excelSerialToDate(serial)` | Excel serial number → `Date`. |
| `isDate(value)` | Type guard for valid `Date` objects. |
| `calculateColumnWidths(data)` | Compute optimal column widths. |

### Types

<details>
<summary>Type definitions</summary>

```typescript
// Values accepted in a cell. Dates become Excel serials; strings starting with "=" are formulas.
type CellValue = string | number | boolean | Date | null | undefined;

interface ExcelData {
  data: CellValue[][];
  styles?: Record<string, CellStyle>;
  validations?: CellValidation[];
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  hyperlinks?: Hyperlink[];
  options?: SheetOptions;
}

interface SheetOptions {
  name?: string;
  state?: 'visible' | 'hidden' | 'veryHidden'; // at least one sheet must stay visible
  freezePane?: { row?: number; col?: number };
  autoWidth?: boolean;
  columnWidths?: number[];
  autoFilter?: AutoFilter;
}

interface AutoFilter {
  range: string;
}

type Hyperlink = ExternalHyperlink | InternalHyperlink;

interface ExternalHyperlink {
  range: string;
  url: string;
  tooltip?: string;
  display?: string;
}

interface InternalHyperlink {
  range: string;
  location: string;
  tooltip?: string;
  display?: string;
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

/** A data-validation rule. Build one with the `dataValidation` helpers. */
interface CellValidation {
  range: string;
  /** Comma-separated values of an inline list; kept for compatibility, empty for other types. */
  options: string;
  type?: 'list' | 'whole' | 'decimal' | 'textLength' | 'date' | 'time' | 'custom';
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan'
    | 'greaterThanOrEqual' | 'lessThanOrEqual';
  /** A formula or value, for example `"1"`, `"$K$1:$K$3"` or `"ISNUMBER(A1)"`. */
  formula1?: string;
  formula2?: string;
  /** Default: true when writing; the reader reports what the file says. */
  allowBlank?: boolean;
}

// Conditional formatting: a discriminated union on `type`.
type ConditionalFormat =
  | CellValueConditionalFormat
  | ExpressionConditionalFormat
  | ColorScaleConditionalFormat;

type ConditionalFormatOperator =
  | 'greaterThan' | 'greaterThanOrEqual'
  | 'lessThan' | 'lessThanOrEqual'
  | 'equal' | 'notEqual'
  | 'between' | 'notBetween';

interface ConditionalFormatStyle {
  background?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

interface CellValueConditionalFormat {
  type: 'cellValue';
  range: string;
  operator: ConditionalFormatOperator;
  value: number | string;
  value2?: number | string; // for "between" / "notBetween"
  style: ConditionalFormatStyle;
}

interface ExpressionConditionalFormat {
  type: 'expression';
  range: string;
  formula: string;
  style: ConditionalFormatStyle;
}

interface ColorScaleConditionalFormat {
  type: 'colorScale';
  range: string;
  colors: [string, string] | [string, string, string]; // 2- or 3-color scale
}

interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'date' | 'error' | 'empty';
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

</details>

### Advanced / low-level exports

For custom pipelines, `excel-bridge` also exports its building blocks: the functional
`parseExcel`, `createExcelFile`/`createExcelFileBuffer`; the ZIP layer
(`createExcelBlob`, `createExcelBuffer`, `extractExcelFiles`, `validateExcelStructure`); the XML
template generators (`generateSheetXml`, `generateStylesXml`, `generateSheetRelsXml`, …); and
date/validation utilities (`dateToExcelSerial`, `excelSerialToDate`, `EXCEL_LIMITS`,
`validateCellValue`, …). These are lower-level and their shape may change in a minor release. Most
apps only need the entry points above.

## Compatibility

| Environment | Support |
| --- | --- |
| Node.js | `^20.19.0`, `^22.13.0`, or `>=24` (matches `engines`). The test suite runs on Node 20, 22 and 24. |
| Browsers | Modern browsers with ES2022, `File` and `Blob` APIs |
| Module formats | ESM (`import`) and CommonJS (`require`) |
| Bun, Deno | The packed package is smoke-tested on Node 22, Bun 1.x and Deno 2.x in CI |
| Excel | Targets Excel 2016+. CI does not open the output in Excel, LibreOffice or Google Sheets. The reader is tested on files written by ExcelJS, SheetJS and hucre |

### Known limitations

- **Inline strings by default.** Enable a shared-strings table with `new ExcelWriter({ sharedStrings: true })` for smaller files with lots of repeated text. The streaming writer ignores that option and always writes inline strings.
- **Formulas recalculate on open.** Formula cells are written without a cached value; Excel computes them on load (`fullCalcOnLoad`). A reader that does not calculate shows them empty: SheetJS with default options skips the cell and openpyxl with `data_only` returns `None`.
- **Strings starting with `=` are formulas.** There is no way to write such text literally yet. Do not pass user-controlled text that starts with `=` to the writers.
- **Date cells cannot be styled.** A `Date` is always written with the built-in date format, and the reader returns no style for date-formatted cells.
- **Dates are local wall-clock values.** `new Date(2024, 0, 15)` is written as 15 January whatever the time zone, with the built-in short date format (no time of day, locale dependent). A UTC-midnight `Date` such as `new Date('2024-01-15')` falls on the previous day in negative offsets.
- **One box border.** `border: true` draws the same border on all four sides. There are no per-side borders, row heights or hidden rows and columns.
- **AutoFilter ranges only.** Filter criteria and sort state are not written or read.
- **Hyperlink schemes.** The writer accepts `http:`, `https:`, `mailto:` and locations inside the workbook.
- **Reader limits.** The reader rejects cell and row references outside Excel's grid (`XFD1048576`), clamps `<col>` ranges to 16,384 columns and throws when a workbook needs more than 5,000,000 empty cells of padding to keep rows rectangular. It still holds the whole file in memory, so cap the upload size before parsing untrusted files.
- **Not read correctly yet:**
  - the 1904 date system
  - shared-formula followers (formula `""`)
  - an ISO-8601 `t="d"` date cell (read as the number 2024)
  - character escapes such as `_x000D_`
  - split panes (their twips are read as a freeze pane of that many rows and columns)
  - XML written with namespace prefixes (an empty sheet, or an `Invalid Excel file structure` error)
  - rows or cells without an `r` attribute (`rowIndex` is `NaN`, a cell without `r` is lost)

Upgrading from an earlier version? See [UPGRADING.md](./UPGRADING.md).

## Security

Report vulnerabilities privately through [GitHub security advisories](https://github.com/KevinArce98/excel-bridge/security/advisories/new)
(see [SECURITY.md](./SECURITY.md)). The reader is not hardened for untrusted files beyond the
limits listed above.

## Contributing

Issues and pull requests are welcome. To work on the library locally:

```bash
pnpm install
pnpm run build          # bundle ESM + CJS + types
pnpm run test           # run the test suite (watch)
pnpm run test:run       # run once (CI / pre-publish)
pnpm run lint           # ESLint
pnpm run format:check   # Prettier
pnpm run check:package  # publint and arethetypeswrong on the built package
pnpm run size:check     # the README bundle size tables against the build
pnpm run smoke:pack     # install the packed tarball and run it on Node, Bun and Deno
```

Commits follow [Conventional Commits](https://www.conventionalcommits.org/); releases are
published automatically by semantic-release. See the [CHANGELOG](./CHANGELOG.md) for release notes.

## License

[MIT](./LICENSE) © [Kevin Arias](https://github.com/KevinArce98)

Microsoft Excel is a trademark of Microsoft. This project is not affiliated with Microsoft.
