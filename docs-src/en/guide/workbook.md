---
title: Edit a workbook
description: Build a workbook with Workbook, or load one, change cells, styles and layout, and save it back.
group: Guide
groupOrder: 2
order: 7
---

# Edit a workbook

`Workbook` builds a workbook, or loads one, edits it and saves it back, without rebuilding sheet data by hand. It rebuilds the file from what the library models, so it suits files this library wrote and simple files from other tools.

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

Load with `Workbook.fromBuffer(buffer)`, or `await Workbook.fromFile(file)` in the browser. Both take the same [reader limits](../errors/) as `ExcelReader`. Save with `toBuffer()` or `toBlob()`.

## Methods

Rows and columns are zero-based, and sheets are addressed by name.

| Area | Methods |
| --- | --- |
| Sheets | `addSheet`, `renameSheet`, `removeSheet`, `getSheetNames`, `getSheetData`, `getSheetState`, `setSheetState` |
| Cells | `getCellValue`, `setCellValue`, `getCellStyle`, `setCellStyle` |
| Layout | `setMergeCells`, `setFreezePane`, `setColumnWidths`, `setAutoWidth`, `setRowHeight`/`getRowHeight`, `setRowHidden`/`isRowHidden`, `setColumnHidden`/`isColumnHidden` |
| Rules and links | `addValidation`, `addConditionalFormat`, `setAutoFilter`, `getAutoFilter`, `removeAutoFilter`, `setHyperlink`, `getHyperlinks`, `removeHyperlink` |
| Document | `getMetadata`, `setMetadata` |
| Output | `toBuffer`, `toBlob` |

## What you get back from a loaded file

`getCellValue` and `getSheetData` return the same values you write. A loaded formula comes back as `{ formula, result? }` with the value stored in the file, and a loaded error as `{ error }`. Text stays text, even when it starts with `=`.

> [!WARNING]
> A formula's `result` is the value the file stored. It does not change when you edit the cells the formula uses, so it can go stale until Excel recalculates on open. Set the cell to `{ formula }` without a `result` to drop it.

`getSheetData(name)` has holes where the file has blank rows: `length` is the last row plus one, `for...of` yields `undefined` for a hole and `JSON.stringify` writes `null`. A loaded sheet whose name Excel would reject, for example one over 31 characters, loads, but saving throws until you give it a valid name with `renameSheet`.

## What a round trip keeps

`Workbook.fromBuffer` and `fromFile` restore:

- data at its row and column, styles, merges, freeze panes and column widths
- per-side borders, row heights, hidden rows and hidden columns
- data validations, conditional formatting rules, autofilter ranges, hyperlinks and sheet visibility
- formulas with their stored results, and the seven classic error values

## What it changes or drops

- Images, charts, comments, Excel tables, pivot tables, defined names, print setup, themes and macros.
- Filter criteria and sort state set in Excel. Rows hidden by a filter stay hidden, and `removeAutoFilter` shows the rows under the removed range.
- Links the writer does not accept (anything but `http:`, `https:`, `mailto:` or a location in the workbook), and validations of a type the reader does not know. `ExcelReader` still returns the links.
- Validation input and error messages, the error style (stop, warning, information) and the "show message" switches. Every saved rule is written to show both messages and to reject bad entries.
- Error text other than the seven classic errors, which is saved as text.
- Locale-dependent built-in date and time formats, which are written back as explicit en-US codes. The built-in short date (id 14) stays the short date.
- Theme and indexed colours, and built-in number formats such as `0%`. Only RGB colours, custom number formats and the built-in date and time formats are read, and theme or indexed border colours become black.
- Diagonal borders, rich text (flattened) and split panes (they become freeze panes).
- Conditional formatting rules other than `cellIs`, `expression` and `colorScale`.
- Sheet defaults (`<sheetFormatPr>`: default row height and column width, outline levels). A collapsed outline group stays hidden but loses its buttons.
- Shared-formula followers, which are read without their formula.
- A stored formula whose text starts with `=`. Excel does not write one, but a file made by 1.x from a string such as `=== Summary ===` has one. `Workbook` saves it without that first `=`, and a formula that is only `=` makes `toBuffer` throw.
- A stored formula whose text starts with `=`. Excel does not write one, but a file made by 1.x from a string such as `=== Summary ===` has one. `Workbook` saves it without that first `=`, and a formula that is only `=` makes `toBuffer` throw.
- Dates in the 1904 date system, which are read incorrectly.

A column stored with width 0 stays 0 wide when you unhide it with `setColumnHidden`. Set its width as well.

> [!LIMIT]
> If you need to edit a file and keep everything you did not touch, such as a template with charts or macros, `Workbook` is the wrong tool: it rebuilds the file from its model. hucre's `openXlsx` and `saveXlsx` keep the parts they do not model.
