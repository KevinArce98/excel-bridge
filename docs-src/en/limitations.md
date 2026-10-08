---
title: Limits and alternatives
description: What excel-bridge does not do, the cases it reads wrong, where it has been tested, and when another library is the better choice.
group: Reference
groupOrder: 3
order: 3
---

# Limits and alternatives

excel-bridge writes reports and reads them back. It does not cover everything, and this page says where it stops.

## When another library fits better

- **Images, charts, comments, Excel tables, pivot tables, print setup, rich text, diagonal borders.** None of these are supported. ExcelJS and hucre cover many of them.
- **Editing a file and keeping everything you did not touch**, such as templates with charts or macros. `Workbook` rebuilds the file from what it models, so anything else is dropped. hucre's `openXlsx` and `saveXlsx` keep the parts they do not model.
- **Reading very large files without loading them whole.** `ExcelReader` has no streaming mode. hucre has `streamXlsxRows`.
- **CSV, ODS or legacy `.xls` and `.xlsb`.** Use hucre or SheetJS.
- **Schema-validated imports** that map rows to typed objects with error lists. Use read-excel-file.

## Where it has been tested

| Environment | Support |
| --- | --- |
| Node.js | `^20.19.0`, `^22.13.0` or `>=24`, matching `engines`. The test suite runs on Node 20, 22 and 24. |
| Browsers | Not run in CI. The code targets ES2022 and needs `File` and `Blob`. `downloadXlsx` is tested against a stubbed `document`. |
| Module formats | ESM (`import`) and CommonJS (`require`). |
| Bun and Deno | The packed package is smoke-tested on Node 22, Bun 1.x and Deno 2.x in CI. |
| Excel | Written for Excel 2016 and later, but CI does not open the output in Excel, LibreOffice or Google Sheets. The reader is tested on files written by ExcelJS, SheetJS and hucre. |

## Writing

- **Formulas recalculate on open.** The workbook asks Excel to recalculate every formula when it loads. A formula without a `result` carries no stored value, so a reader that does not calculate shows it empty: SheetJS with default options skips the cell, and openpyxl with `data_only` returns `None`.
- **Dates are local wall-clock values.** `new Date(2024, 0, 15)` is written as 15 January in every time zone, with the built-in short date format unless the cell's style has a date `numberFormat`. A UTC-midnight date falls on the previous day in negative offsets, and a local time inside a daylight-saving gap moves forward.
- **Borders cover the four sides**, with 13 line styles and RGB colours. There are no diagonal borders, and no borders in conditional formats.
- **Layout defaults.** A row without a height uses Excel's default. Sheet defaults (default row height and column width, outline levels) are not written or kept.
- **AutoFilter ranges only.** Filter criteria and sort state are not written or read.
- **Hyperlink schemes.** The writer accepts `http:`, `https:`, `mailto:` and locations inside the workbook.
- **The streaming writer** does not support `autoWidth`, `validations`, `conditionalFormats`, `sharedStrings` or a sheet `state`, and it always writes strings inline.

## Reading

The reader is not hardened for untrusted files beyond its [limits](../guide/errors/). It holds the whole file in memory.

These cases are not read correctly yet:

- the 1904 date system
- an ISO-8601 `t="d"` date cell, which is read as a year near 1905
- character escapes such as `_x000D_`, which stay in the text
- split panes, whose twips are read as a freeze pane of that many rows and columns
- a shared-formula follower, which is read as its cached value without a formula

## What it does not claim

- It is not the smallest on every row of the comparison. See [Performance and size](../performance/).
- It is not the fastest: hucre writes the benchmark workload faster (about 390 ms against 560 ms).
- Nothing here says a file was tested in Excel, LibreOffice or Google Sheets, because CI does not open one.
