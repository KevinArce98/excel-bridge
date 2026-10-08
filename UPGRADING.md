# Upgrading

Notes for changes that can affect code written against an earlier version. The release notes list
what changed; this file says what to do about it.

## From 1.4 to 1.5

### Calls that now throw

The writers used to produce a file for these inputs. A file Excel repairs, or that other readers
reject, is worse than an error at the call.

| Input | What to do |
| --- | --- |
| A sheet name over 31 characters, with `\ / ? * [ ] :`, a control character or an apostrophe at either end, or two names equal ignoring case | Shorten or rename it. |
| A workbook with no sheets, or with every sheet hidden | Add a sheet, or leave one visible. |
| A colour that is not `#RGB`, `#RRGGBB` or `#AARRGGBB` (`red`, `rgb(…)`) | Convert it to hex. |
| `NaN`, `Infinity` or an invalid `Date` in a cell | Replace it before writing, for example `Number.isFinite(value) ? value : null`. The error names the cell. |

This applies to `ExcelWriter`, `createExcelWorkbookStream`, `Workbook.toBuffer`/`toBlob`, and to
`Workbook.addSheet`, which now also compares names ignoring case.

A workbook loaded with `Workbook.fromBuffer` keeps the names it had. If one is longer than 31
characters, loading works and saving throws; call `workbook.renameSheet(from, to)` first. Formulas
that mention the old name are not rewritten.

The reader throws for a cell or row outside Excel's grid (past column `XFD` or row 1,048,576) and
for a workbook that needs more than 5,000,000 empty cells of padding to keep rows rectangular. Both
used to risk exhausting memory.

### Values the reader returns

- **Error cells** (`#DIV/0!`) used to read as the number `NaN`. They now read as `type: 'error'`
  with the text as the value, so a `switch` over `ParsedCell['type']` needs a case for `'error'`.
  A number that is not finite (`1e999`) reads as the error `#NUM!`.
- **No numeric conversion.** Attribute and text values are returned as they are written. A number
  format `0.00` used to read as `0`, and a sheet name or title `007` as `7`. If you relied on a
  number there, convert it yourself.
- **`ParsedSheet.validations`** holds full rules (`type`, `operator`, `formula1`, `formula2`,
  `allowBlank`) instead of `{ range, options }`. `options` still holds the values of an inline list.
  For other types it is now the first formula with its quotes intact (they used to be stripped), and
  rules of type `none` are no longer returned.
- **`ParsedSheet.state`** is `'hidden'` or `'veryHidden'` for hidden sheets and absent otherwise.
- **`DataValidationType`** gains `'time'` and `'custom'`; an exhaustive `switch` needs both.

`ParsedSheet.data` is unchanged: the rows present in the file, in order. Use `cell.rowIndex` for the
position.

### Workbook

- **Rows stay where they were.** Loading a file with blank rows and saving it used to move every
  later row up while styles and merges stayed put. Rows now keep their position.
- **`getSheetData(name)` can have holes.** For a file with blank rows it is sparse: `length` is the
  last row plus one, `for...of` yields `undefined` for a missing row and `JSON.stringify` writes
  `null`. Use index loops with `?.`, or `forEach`, which skips holes. `getCellValue(sheet, row, col)`
  now agrees with the row in the file.
- **Validations, number formats and hidden sheets survive a save.** Validation messages and the
  error style do not; see the round trip note in the README.
- **New:** `renameSheet`, `getSheetState`, `setSheetState`.

### Packaging

- The `exports` map gives ESM importers `index.d.mts` and CommonJS importers `index.d.ts`, exports
  `./package.json`, and the package declares `"sideEffects": false` and `"type": "commonjs"`.
- The low-level exports (`generate*Xml`, the zip helpers) are no longer called stable.

### Output and size

- The streaming writer deflates every zip entry, so SheetJS 0.18.5 and Java's `ZipInputStream` can
  read its output.
- `ExcelWriter` grows from 12.0 to 12.5 KB min+gzip and `ExcelReader` from 27.6 to 28.2 KB.
