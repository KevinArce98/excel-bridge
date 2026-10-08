# Upgrading

Notes for changes that can affect code written against an earlier version. The release notes list
what changed; this file says what to do about it.

## From 1.5 to 1.6

Everything you could write in 1.5 still writes the same file, apart from the output notes below.
New features are additive. Four kinds of change can affect code: TypeScript types that widen, a few
calls that now throw, output and reader values that differ, and `Workbook` round trips that keep more.

### Types that widen (TypeScript only)

The runtime values of these do not change, except `ParsedSheet.styles[...].border` (see "Values the
reader returns"), but code that narrows the old types can stop compiling.

- **`CellValue`** also covers `FormulaCell`, `TextCell` and `ErrorCell`. Code that reads values
  (`getCellValue`, `getSheetData`) and checks `typeof value === 'object'` to find a `Date` needs
  `value instanceof Date`. Assigning `getSheetData()` to an array of the old union (`string | number | boolean | Date | null |
  undefined`) needs `CellValue`.
- **Array literals of cell objects** such as `[[{ error: '#N/A' }]]` infer `string` for the error.
  Annotate the array as `CellValue[][]`, or write `{ error: '#N/A' as const }`.
- **`CellStyle.border`** is `CellBorder`: `boolean`, a line style, `{ style, color? }` or per-side
  entries. `ParsedSheet.styles[...].border` is `true` for the plain thin box and an object with one
  entry per side otherwise, so `const border: boolean | undefined = style.border` no longer
  compiles. Truthiness checks still work.
- **`ParsedSheet`, `SheetOptions` and `StreamingSheetInput`** gain `rowHeights`, `hiddenRows` and
  `hiddenColumns`. An exhaustive check over their keys needs the new ones.
- **`generateColsXml(widths, layout?)`** takes a layout object as its second argument and ignores a
  number, so passing it straight to `Array.prototype.map` still runs but no longer type-checks.
  A subclass of `StyleManager` that declares its own private `registerStyle`, `styleIds` or
  `internXf` conflicts with the new private members.

### Calls that now throw

| Input | What to do |
| --- | --- |
| `{ formula: '' }` (or `{ formula: '=' }`) | Pass a formula. The string shorthand `'='` is unchanged. |
| `{ error: '#SPILL!' }` or any error outside the seven classic ones | Use `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!` or `#N/A`, or write the text with `{ text }`. |
| A row height of 0 or above 409.5, or a hidden row or column index that is not a whole number (rows 0 to 1,048,575, columns 0 to 16,383) | Fix the value. |
| A border line style other than the 13 names, a colour that is not hex, or a truthy value that is not a border (`1`, `'true'`, `{ left: true }`) | Use `true`, a line style, `{ style, color? }` or per-side entries. |
| A border object that mixes `color` with side keys (`{ left: 'thin', color: '#000000' }`) | Put the colour inside each side. TypeScript also rejects `{ left: 'thin', style: 'thin' }`; at runtime its side keys are ignored. |

An object with `formula`, `text` or `error` keys used to be written as the text `[object Object]`, and
`{ formula: 1 }` or `{ text: 1 }` now throw `Cell A1 needs a string formula or text`.
`{}` and `[]` as a `border` (not typed, JavaScript only) used to draw a box and now mean no border.

### Output

- **Equal styles share one entry.** Redundant styles (explicit `false` flags, `#f00` next to
  `#FF0000`, `fontName: 'Calibri'` spelled out, equal conditional-format styles) now collapse, so the
  counts in `styles.xml` and the `s=` indexes can shift. `ExcelReader` returns the same values and
  styles. `border: false` (or `null`, `0`, `''`) no longer adds a second empty border.
- **`generateStylesXml()` without an argument** returns what `generateStylesXml(new StyleManager())`
  returns, which is the `styles.xml` of an `ExcelWriter` file with no styles.
- **A style on a `Date` cell is written.** Bold, fill, border, alignment and a date `numberFormat`
  used to be ignored. A `numberFormat` that is not a date format is still ignored on a `Date` cell,
  and a hyperlink over a `Date` cell now gives it the hyperlink font.
- **A hidden column without a width** is written with width `9.140625`, the default column width in Excel files.
- **Size:** `ExcelWriter` goes from 12.5 to 12.9 KB min+gzip, the streaming writer from 11.9 to
  12.3 KB, `ExcelReader` from 28.2 to 28.7 KB and `Workbook` from 39.9 to 41.0 KB.

### Values the reader returns

- **Cells with a date number format report their style** when it also has a custom number format,
  font, fill, border or alignment, whatever the cell type (a string or an empty cell can carry one),
  and `numberFormat` carries the code for the built-in ids 15 to 22 and 45 to 47. A plain short-date
  cell still reports none, and `Workbook.getCellStyle` changes the same way.
- **`border` on a read style** used to be `true` for any border. It is now `true` only for a thin
  black line on all four sides, and otherwise an object with a `{ style, color? }` entry for each
  side that has a line, for example `{ left: { style: 'thin' } }`. Replace `style.border === true`
  with a truthiness check.
- **An empty numeric value** (`<v></v>`) reads as an empty cell instead of `#NUM!`. A `<col>`
  without a width no longer reads as `NaN`, and a border with `style="none"` no longer reads as a
  border.
- **`rowHeights`, `hiddenRows` and `hiddenColumns`** appear on `ParsedSheet` only for files that
  have them.

### Workbook

- **A round trip now keeps** per-side borders, row heights, hidden rows and hidden columns, date
  cell styles and custom date formats, the seven classic error cells, and text that starts with
  `=`. A file that relied on a save un-hiding rows or columns now keeps them hidden.
- **Loaded errors and `=` text keep their kind when saved.** `getCellValue` and `getSheetData`
  still return the same strings (`'#N/A'`, `'=== Summary ==='`), so code that only reads values is
  unaffected. Writing those strings back with `setCellValue` or `addSheet` stores a formula and
  text, so wrap them in `{ text }` and `{ error }`. Before, a loaded `=== Summary ===` was saved as a
  formula.
- **`removeAutoFilter` shows the rows under the removed range.** Rows hidden by a filter stay
  hidden when you load and save, because the criteria are not kept.
- **Cached formula results are not kept.** Formulas still load as `'=…'` strings and save without a
  result.
- **New:** `setRowHeight`, `getRowHeight`, `setRowHidden`, `isRowHidden`, `setColumnHidden` and
  `isColumnHidden`.

### New and recommended

`{ formula, result? }`, `{ text }` and `{ error }` cells; per-side borders with 13 line styles and
RGB colours; styles on `Date` cells; `rowHeights`, `hiddenRows` and `hiddenColumns` on
`ExcelWriter`, the streaming writer and `Workbook`. Prefer `{ formula }` and `{ text }` in new code:
the `'='` string shorthand stays a formula for all of 1.x, and `{ text }` is the way to store text
that starts with `=`. The exported types `Font`, `Fill`, `Border`, `ExcelStyle` and `CellAlignment`
are no longer used by the library and will be removed in 2.0.

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
