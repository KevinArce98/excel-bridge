---
title: Getting started
description: Install excel-bridge, write a styled .xlsx and read it back, in the browser and Node.js.
group: Start
groupOrder: 1
order: 1
---

# Getting started

Write a styled spreadsheet and read it back, in the browser or Node.js, from one small package. Import a class and your bundler ships only that part: the writer adds 12.9 KB min+gzip.

## Install

```bash tab="npm"
npm install excel-bridge
```

```bash tab="pnpm"
pnpm add excel-bridge
```

```bash tab="yarn"
yarn add excel-bridge
```

```bash tab="bun"
bun add excel-bridge
```

excel-bridge needs Node.js 20.19, 22.13 or 24 and later, or a modern browser. It is written in TypeScript and ships ESM and CommonJS builds with types.

## Write a workbook

Each sheet is an object with `data` and, optionally, `styles` and `options`. Styles are keyed by `"<row>-<col>"` with zero-based indexes.

```ts title="write.ts"
import fs from 'node:fs';
import { ExcelWriter } from 'excel-bridge';

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

const writer = new ExcelWriter();

fs.writeFileSync('people.xlsx', writer.createWorkbookBuffer([sheet]));
```

In the browser, `createWorkbook([sheet])` returns a `Blob`. Pass it to `downloadXlsx(blob, 'people')` to start the download.

> [!NOTE]
> `createWorkbookBuffer` returns a plain `Uint8Array`, not a `Buffer`. Express 4 needs `res.send(Buffer.from(bytes))`.

## Read it back

```ts title="read.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('people.xlsx'));

workbook.sheets[0].data.forEach(row => {
  row.forEach(cell => console.log(cell.coordinate, cell.type, cell.value));
});
```

Every cell carries its `type`. Dates come back as `Date`, and a formula cell carries its formula in `formula`. In the browser, `await reader.parseFromFile(file)` takes the `File` from an `<input type="file">`.

## Edit a file

`Workbook` loads a file, lets you change it and saves it back.

```ts title="edit.ts"
import fs from 'node:fs';
import { Workbook } from 'excel-bridge';

const workbook = Workbook.fromBuffer(fs.readFileSync('people.xlsx'));
workbook.setCellValue('People', 1, 1, 26);
fs.writeFileSync('people.xlsx', workbook.toBuffer());
```

> [!LIMIT]
> `Workbook` rebuilds the file from what it models, so images, charts and comments of a loaded file are dropped. See [Edit a workbook](../guide/workbook/).

## Where to go next

- [Write values and formulas](../guide/values/): what a cell can hold, and how to write a formula.
- [Style cells](../guide/styling/) and [Lay out sheets](../guide/layout/): borders, widths, merged cells, filters.
- [Read a workbook](../guide/reading/) and [Rows as objects](../guide/objects/): typed cells and typed rows.
- [Stream big exports](../guide/streaming/): million-row files and file delivery.
- [API reference](../api/) and [Upgrading](../upgrading/).
