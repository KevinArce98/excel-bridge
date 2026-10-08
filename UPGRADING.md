# Upgrading

Notes for changes that can affect code written against an earlier version. The release notes list
what changed; this file says what to do about it.

## From 1.x to 2.0

Version 2.0 changes what a plain string means, what the reader and `Workbook` return, and what the package exports. Most code needs a search and a few edits. Do the first step on 1.6, before you upgrade.

### Before you upgrade: write formulas as objects

In 1.x a string that starts with `=` was written as a formula. In 2.0 it is text. The change fails silently: a cell that held `=SUM(A1:A3)` as a formula now shows those characters as text.

| You write | 1.x | 2.0 |
| --- | --- | --- |
| `'=SUM(A1:A3)'` | a formula | the text `=SUM(A1:A3)` |
| `{ formula: 'SUM(A1:A3)' }` | a formula | a formula, same bytes |
| `{ formula: 'x', result: 4 }` | a formula with a stored value | the same |
| `{ text: '=SUM(A1:A3)' }` | text | text, same bytes as the plain string |

`{ formula }` and `{ text }` already exist in 1.6 and mean the same there, so you can migrate first and upgrade after. The files do not change.

Find the strings with this ESLint rule:

```js
{
  rules: {
    'no-restricted-syntax': [
      'error',
      {
        selector: 'Literal[value=/^=/]:not(Property[key.name="text"] > Literal)',
        message: "A string that starts with '=' is text in excel-bridge 2.0. Use { formula: '...' } for a formula.",
      },
      {
        selector: 'TemplateLiteral > TemplateElement.quasis:first-child[value.raw=/^=/]',
        message: "A string that starts with '=' is text in excel-bridge 2.0. Use { formula: '...' } for a formula.",
      },
    ],
  },
}
```

If formulas reach you as strings from a config file, a database or an API, convert them at the boundary. This turns formula injection back on for that data, so do it only for data you trust.

```ts
import type { CellValue } from 'excel-bridge';

const asFormulas = (rows: CellValue[][]): CellValue[][] =>
  rows.map(row =>
    row.map(cell =>
      typeof cell === 'string' && cell.startsWith('=') ? { formula: cell.slice(1) } : cell
    )
  );
```

The benefit is that exporting user data can no longer create a formula by accident. Text such as `=== Summary ===` needs no wrapper any more.

### Writing

- **`CellValidation.options` is gone**, in the rules you write and in `ParsedSheet.validations`. A list rule uses `formula1`. `dataValidation.list(range, values)` is unchanged for callers. A hand-written list needs `formula1: '"a,b"'`, and a list without `formula1` throws `Validation at A2:A5 needs formula1`. To read the values of an inline list: `rule.formula1?.match(/^"(.*)"$/s)?.[1].replace(/""/g, '"').split(',')`.
- **`ExcelWriter.addValidation`, `addStyle`, `createSimple` and `createSimpleBuffer` are removed.** Set `data[data.length - 1].validations` and `.styles` yourself, or use `dataValidation`, and use `createExcelFile` or `createExcelFileBuffer` for one sheet.
- **`SheetLayout` is complete.** It now holds `freezePane` and `columnWidths` as well as `rowHeights`, `hiddenRows` and `hiddenColumns`. `calculateColumnWidths` takes `CellValue[][]`.

### Reading

- **`ParsedCell` is a union on `type`.** `cell.value` is no longer `any`. Narrow on `cell.type` and `value` has the right type. An `'error'` cell has `ExcelErrorValue | string`. A formula cell has the type and value of its stored result, and `formula` is absent or non-empty.
- **`ParsedSheet.data` is indexed by row index.** `data[4]` is row 5. A row that is not in the file is a hole, `data.length` is the last row plus one, and an empty `<row>` element is a hole too, so a round trip no longer adds rows. Use `forEach`, `Object.values` or `flat()`. A `for...of` yields `undefined` for a hole, spreading the array turns each hole into `undefined`, and `JSON.stringify` writes `null` for each one: a file whose only row is at index 1,000,000 produces over 5 million characters (5,000,090 for one numeric cell).
- **Rows and cells without an `r` attribute** get the position after the previous one, so `rowIndex` is never `NaN`. Duplicate rows merge, and the later cell wins.
- **A shared-formula follower has no `formula`.** It reads as its stored value, which `Workbook` saves as a plain value.
- **`ParsedCellStyle` is replaced by `CellStyle`.** A style read from a file can be written back as it is. `CellStyle` takes an optional type argument for the border, with a default, so existing uses keep working.

### Workbook

- **`getCellValue` and `getSheetData` return what they load.** A loaded formula is `{ formula, result? }` and a loaded error is `{ error }`. Text stays text, including text that starts with `=`. Code that compared a value with `'#N/A'` or printed `String(value)` for a formula cell now sees an object. `typeof value === 'object' && value !== null && 'formula' in value` narrows to a formula, and `'error' in value` to an error.
- **Stored formula results are kept.** 1.6 dropped them on save. A result does not change when you edit the cells it depends on, so it can go stale until Excel recalculates on open. To drop one, set the cell to `{ formula: cell.formula }`.
- **`splice` and `unshift` on `getSheetData` rows are safe.** Nothing is tracked by position any more.

### Errors and limits

- **Everything the library raises on purpose is an `ExcelBridgeError`** with a `code` (`INVALID_INPUT`, `INVALID_FILE`, `LIMIT_EXCEEDED` or `UNSUPPORTED`). Messages are unchanged, apart from the cell budget one below. `error.name` is now `ExcelBridgeError`, so code that compared it with `'Error'` should use `isExcelBridgeError(error)` or `error instanceof Error`. A `parseFromBuffer` error keeps the `Failed to parse Excel file:` prefix and puts the original error on `cause`.
- **`new ExcelReader()` has limits.** It refuses files above `maxCells: 5_000_000` (now every cell, not only padding), `maxPartBytes: 268_435_456` and `maxTotalBytes: 536_870_912`, with `LIMIT_EXCEEDED`. A part that declares a size over its limit is refused before it is inflated. Pass `Infinity` for each to remove the caps. The message `Workbook pads more than 5000000 empty cells to keep rows rectangular` is now `Workbook has at least N cells, counting the empty cells that pad rows, over the limit maxCells of 5000000`.
- **The XML reader is stricter.** A mismatched or unclosed tag, an unquoted attribute and a document that ends early now throw `INVALID_FILE`, where 1.x read what it could. A part with a `DOCTYPE` or in UTF-16 is rejected. A malformed `docProps` part is skipped, so its metadata is empty. A numeric character reference such as `&#233;` is decoded.
- **Files with prefixed XML namespaces read correctly.** A sheet written as `<x:worksheet>` no longer reads as empty.
- **`t="str"` cells with `xml:space="preserve"`**, as SheetJS writes them, read as their text instead of `[object Object]`.

### Removed exports

| Removed | Use instead |
| --- | --- |
| `StyleManager` | Nothing: it only served the XML generators. |
| `generateSheetXml`, `generateSharedStringsXml`, `generateStylesXml`, `generateContentTypesXml`, `generateWorkbookXml`, `generateWorkbookRelsXml`, `generateRootRelsXml`, `generateCorePropsXml`, `generateAppPropsXml`, `generateSheetRelsXml`, `generateColsXml` | `ExcelWriter` or `createExcelWorkbookStream` for whole files. The parts cannot be used alone. |
| `createExcelBlob`, `createExcelBuffer`, `extractExcelFiles`, `validateExcelStructure` | `fflate` directly (`zipSync`, `unzipSync`, `strToU8`, `strFromU8`). It is already a dependency. |
| `XML_NS`, `CONTENT_TYPES`, `RELATIONSHIP_TYPES`, `CELL_TYPES` | Nothing: the library hard-codes them. |
| `isDateNumFmtId`, `isDateFormatCode`, `validateRowIndex`, `validateColIndex`, `validateCellValue` | Nothing: reader and writer internals. |
| Types `ParsedCellStyle`, `Font`, `Fill`, `Border`, `ExcelStyle`, `CellAlignment`, `ExcelFiles`, `SheetGenerationOptions`, `DefinedName` | `CellStyle` for styles. The others have no replacement. |

`parseExcel`, `createExcelFile`, `createExcelFileBuffer` and `ExcelBridge` stay. `isExcelError` is new.

### New

- `sheetToObjects`, `objectsToSheet` and `objectsToStreamingSheet`: rows as typed objects.
- `downloadXlsx`, `toReadableStream`, `xlsxResponse` and `XLSX_CONTENT_TYPE`: send a file to a browser or from a server.
- `ExcelBridgeError`, `isExcelBridgeError` and `isExcelError`.
- Reader options on `ExcelReader`, `parseExcel`, `ExcelBridge.read`, `ExcelBridge.readFromFile`, `Workbook.fromBuffer` and `Workbook.fromFile`.

### Size and speed

Bundle size, min+gzip, measured with the same esbuild call on both versions:

| Entry | 1.6.0 | 2.0 |
| --- | ---: | ---: |
| `createExcelWorkbookStream` | 12.3 KB | 12.3 KB |
| `ExcelWriter` | 12.9 KB | 12.8 KB |
| `ExcelReader` | 28.8 KB | 9.5 KB |
| `Workbook` | 41.0 KB | 21.5 KB |
| `ExcelBridge` | 41.2 KB | 21.6 KB |
| Everything | 43.7 KB | 25.1 KB |

Reading a 50,000 × 10 file takes about 0.27 s instead of 1.5 s, and the peak memory of the benchmark process drops from 859 to 451 MiB (Node 24.19, Apple M4). The reader is faster because it parses the XML with its own tokenizer, and the package now depends on `fflate` alone. Writing is as fast as before.

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
  unaffected. Writing those strings back with `setCellValue` or `addSheet` stores `'=== Summary ==='`
  as a formula and `'#N/A'` as text, so wrap them in `{ text }` and `{ error }`. Before, a loaded `=== Summary ===` was saved as a
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
