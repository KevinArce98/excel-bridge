---
title: API reference
description: Every export of excel-bridge with its signature, grouped by what you use it for.
group: Reference
groupOrder: 3
order: 1
---

# API reference

Everything comes from the package root: `import { ExcelWriter } from 'excel-bridge'`. Import the class or function you need and your bundler ships only that part.

## Write

| Export | Signature | Notes |
| --- | --- | --- |
| `ExcelWriter` | `new ExcelWriter(options?)` | `createWorkbook(sheets): Blob` for the browser, `createWorkbookBuffer(sheets): Uint8Array` for Node.js. |
| `createExcelFile` | `(data, options?) => Blob` | One sheet from a 2D array. |
| `createExcelFileBuffer` | `(data, options?) => Uint8Array` | The same, as bytes. |
| `createExcelWorkbookStream` | `(sheets, options?) => AsyncGenerator<Uint8Array>` | Chunks of an `.xlsx` for large exports. See [Stream big exports](../guide/streaming/). |
| `streamToBuffer` | `(stream) => Promise<Uint8Array>` | Collect a stream into one buffer. |
| `dataValidation` | `list`, `wholeNumber`, `decimal`, `textLength`, `dateBetween` | Builders that return a `CellValidation`. |
| `hyperlink` | `url`, `email`, `internal` | Builders that return a hyperlink. |

`ExcelWriterOptions` is `{ creator?, title?, subject?, sharedStrings? }`. Each sheet is an `ExcelData`: `data` plus optional `styles`, `validations`, `mergeCells`, `conditionalFormats`, `hyperlinks` and `options`.

## Read

| Export | Signature | Notes |
| --- | --- | --- |
| `ExcelReader` | `new ExcelReader(options?)` | `parseFromBuffer(bytes): ParsedWorkbook`, `parseFromFile(file): Promise<ParsedWorkbook>`. |
| `parseExcel` | `(bytes, options?) => ParsedWorkbook` | The same as `parseFromBuffer`, as a function. |
| `sheetToObjects` | `(sheet, options?) => Row[]` | A sheet as typed objects. See [Rows as objects](../guide/objects/). |
| `isExcelErrorValue` | `(value) => value is ExcelErrorValue` | True for the seven classic error values. |

`ExcelReaderOptions` is `{ maxCells?, maxPartBytes?, maxTotalBytes?, maxSheets? }`. See [Handle errors and limits](../guide/errors/).

## Edit

`Workbook` loads, edits and saves. `Workbook.create()`, `Workbook.fromBuffer(bytes, options?)` and `Workbook.fromFile(file, options?)` return a workbook; `toBuffer()` and `toBlob()` save it. The methods are listed in [Edit a workbook](../guide/workbook/).

## Objects

| Export | Signature |
| --- | --- |
| `objectsToSheet` | `(rows, columns, options?) => ExcelData` |
| `objectsToStreamingSheet` | `(rows, columns, options?) => StreamingSheetInput` |

## Deliver

| Export | Signature | Notes |
| --- | --- | --- |
| `downloadXlsx` | `(data, filename?) => void` | Browser only. Throws `UNSUPPORTED` without a `document`. |
| `xlsxResponse` | `(body, filename?, init?) => Promise<Response>` | For fetch-style servers. |
| `toReadableStream` | `(chunks) => ReadableStream<Uint8Array>` | Wrap any async iterable of bytes. |
| `XLSX_CONTENT_TYPE` | `string` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`. |

## Errors

| Export | Notes |
| --- | --- |
| `ExcelBridgeError` | Extends `Error`. Has `code` (`ExcelBridgeErrorCode`), an optional `limit` (`ReaderLimitName`) and `cause`. |
| `isExcelBridgeError` | `(error) => error is ExcelBridgeError`. Prefer it to `instanceof`. |

## Helpers

| Export | Signature |
| --- | --- |
| `coordinateToIndex` | `(coordinate) => { row, col }`, so `"A1"` is `{ row: 0, col: 0 }`. |
| `indexToCoordinate` | `(row, col) => string`, so `(0, 0)` is `"A1"`. |
| `dateToExcelSerial` | `(date) => number` |
| `excelSerialToDate` | `(serial) => Date` |
| `isDate` | `(value) => value is Date`, true only for a valid `Date`. |
| `calculateColumnWidths` | `(data) => number[]`, the widths `autoWidth` would use. |
| `EXCEL_LIMITS` | Excel's own limits: rows, columns, cell length, hyperlinks. |

## ExcelBridge

`ExcelBridge` puts `read`, `readFromFile`, `write`, `writeBuffer`, `Reader`, `Writer`, `Workbook`, `coordinateToIndex` and `indexToCoordinate` on one object. It suits scripts and prototypes.

> [!NOTE]
> Bundlers keep an object whole, so even one `ExcelBridge.write` call ships the reader too. In browser code, import the named exports.

## Types

The types you will use most. Every type is exported from the package root.

```ts title="values.d.ts" check=false
type CellValue =
  | string | number | boolean | Date | null | undefined
  | FormulaCell | TextCell | ErrorCell;

interface FormulaCell { formula: string; result?: FormulaResult }
type FormulaResult = string | number | boolean | Date | ErrorCell;
interface TextCell { text: string }
interface ErrorCell { error: ExcelErrorValue }
type ExcelErrorValue = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';
```

```ts title="sheets.d.ts" check=false
interface ExcelData {
  data: CellValue[][];
  styles?: Record<string, CellStyle>;
  validations?: CellValidation[];
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  hyperlinks?: Hyperlink[];
  options?: SheetOptions;
}

interface SheetOptions extends SheetLayout {
  name?: string;
  state?: 'visible' | 'hidden' | 'veryHidden';
  autoWidth?: boolean;
  autoFilter?: AutoFilter;
}

interface SheetLayout {
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  rowHeights?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
}

interface StreamingSheetInput extends SheetLayout {
  name?: string;
  rows: Iterable<CellValue[]> | AsyncIterable<CellValue[]>;
  styles?: Record<string, CellStyle>;
  mergeCells?: string[];
  autoFilter?: AutoFilter;
  hyperlinks?: Hyperlink[];
}
```

```ts title="styles.d.ts" check=false
interface CellStyle<Border extends CellBorder = CellBorder> {
  background?: string;
  border?: Border;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSize?: number;
  fontName?: string;
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
}

type CellBorder = boolean | BorderSide | BorderSides;
type BorderSide = BorderStyleName | BorderLine;
interface BorderLine { style: BorderStyleName; color?: string }
type BorderSides = { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide };
type BorderStyleName =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot'
  | 'slantDashDot';
```

```ts title="parsed.d.ts" check=false
type ParsedCell =
  | ParsedStringCell | ParsedNumberCell | ParsedBooleanCell
  | ParsedDateCell | ParsedErrorCell | ParsedEmptyCell;

interface ParsedCellBase {
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
  formula?: string;
}

interface ParsedNumberCell extends ParsedCellBase { type: 'number'; value: number }
interface ParsedErrorCell extends ParsedCellBase { type: 'error'; value: ExcelErrorValue | (string & {}) }

interface ParsedSheet extends SheetLayout {
  name: string;
  data: ParsedRow[];
  validations: CellValidation[];
  state?: 'hidden' | 'veryHidden';
  styles?: Record<string, CellStyle<ParsedBorder>>;
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  autoFilter?: AutoFilter;
  hyperlinks?: Hyperlink[];
}

interface ParsedWorkbook {
  sheets: ParsedSheet[];
  metadata: { created?: string; modified?: string; creator?: string; title?: string; subject?: string };
}
```

The other cell shapes, `ParsedStringCell`, `ParsedBooleanCell`, `ParsedDateCell` and `ParsedEmptyCell`, follow the table in [Read a workbook](../guide/reading/). `ParsedRow` is an array of `ParsedCell`. `ParsedBorder` is `true` or `BorderLines`, an object with a `BorderLine` entry for each side that has a line.

### All the other types

`AutoFilter`, `BorderLines`, `BorderSideName`, `CellValidation`, `CellValueConditionalFormat`, `ColorScaleConditionalFormat`, `ConditionalFormat`, `ConditionalFormatOperator`, `ConditionalFormatStyle`, `DataValidationOperator`, `DataValidationType`, `ExcelBridgeErrorCode`, `ExcelBridgeErrorOptions`, `ExcelReaderOptions`, `ExcelWriterOptions`, `ExpressionConditionalFormat`, `ExternalHyperlink`, `Hyperlink`, `HyperlinkOptions`, `InternalHyperlink`, `ObjectCellValue`, `ObjectsToSheetOptions`, `ObjectsToStreamingSheetOptions`, `ReadColumn`, `ReaderLimitName`, `SheetState`, `SheetToObjectsOptions`, `StreamingColumn`, `WorkbookMetadata` and `WriteColumn`.
