---
title: Referencia de la API
description: Todas las exportaciones de excel-bridge con su firma, agrupadas según para qué las usas.
group: Referencia
groupOrder: 3
order: 1
---

# Referencia de la API

Todo viene de la raíz del paquete: `import { ExcelWriter } from 'excel-bridge'`. Importa la clase o función que necesites y tu empaquetador incluye solo esa parte.

## Escribir

| Export | Firma | Notas |
| --- | --- | --- |
| `ExcelWriter` | `new ExcelWriter(options?)` | `createWorkbook(sheets): Blob` para el navegador, `createWorkbookBuffer(sheets): Uint8Array` para Node.js. |
| `createExcelFile` | `(data, options?) => Blob` | Una hoja a partir de un arreglo 2D. |
| `createExcelFileBuffer` | `(data, options?) => Uint8Array` | Lo mismo, como bytes. |
| `createExcelWorkbookStream` | `(sheets, options?) => AsyncGenerator<Uint8Array>` | Fragmentos de un `.xlsx` para exportaciones grandes. Consulta [Exporta archivos grandes con streaming](../guide/streaming/). |
| `streamToBuffer` | `(stream) => Promise<Uint8Array>` | Reúne un stream en un solo buffer. |
| `dataValidation` | `list`, `wholeNumber`, `decimal`, `textLength`, `dateBetween` | Constructores que devuelven un `CellValidation`. |
| `hyperlink` | `url`, `email`, `internal` | Constructores que devuelven un hipervínculo. |

`ExcelWriterOptions` es `{ creator?, title?, subject?, sharedStrings? }`. Cada hoja es un `ExcelData`: `data` más, de forma opcional, `styles`, `validations`, `mergeCells`, `conditionalFormats`, `hyperlinks` y `options`.

## Leer

| Export | Firma | Notas |
| --- | --- | --- |
| `ExcelReader` | `new ExcelReader(options?)` | `parseFromBuffer(bytes): ParsedWorkbook`, `parseFromFile(file): Promise<ParsedWorkbook>`. |
| `parseExcel` | `(bytes, options?) => ParsedWorkbook` | Lo mismo que `parseFromBuffer`, como función. |
| `sheetToObjects` | `(sheet, options?) => Row[]` | Una hoja como objetos con tipo. Consulta [Filas como objetos](../guide/objects/). |
| `isExcelError` | `(value) => value is ExcelErrorValue` | Verdadero para los siete valores de error clásicos. |

`ExcelReaderOptions` es `{ maxCells?, maxPartBytes?, maxTotalBytes?, maxSheets? }`. Consulta [Maneja errores y límites](../guide/errors/).

## Editar

`Workbook` carga, edita y guarda. `Workbook.create()`, `Workbook.fromBuffer(bytes, options?)` y `Workbook.fromFile(file, options?)` devuelven un libro; `toBuffer()` y `toBlob()` lo guardan. Los métodos se enumeran en [Edita un libro](../guide/workbook/).

## Objetos

| Export | Firma |
| --- | --- |
| `objectsToSheet` | `(rows, columns, options?) => ExcelData` |
| `objectsToStreamingSheet` | `(rows, columns, options?) => StreamingSheetInput` |

## Entregar

| Export | Firma | Notas |
| --- | --- | --- |
| `downloadXlsx` | `(data, filename?) => void` | Solo en el navegador. Lanza `UNSUPPORTED` sin un `document`. |
| `xlsxResponse` | `(body, filename?, init?) => Promise<Response>` | Para servidores de estilo fetch. |
| `toReadableStream` | `(chunks) => ReadableStream<Uint8Array>` | Envuelve cualquier iterable asíncrono de bytes. |
| `XLSX_CONTENT_TYPE` | `string` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`. |

## Errores

| Export | Notas |
| --- | --- |
| `ExcelBridgeError` | Extiende `Error`. Tiene `code` (`ExcelBridgeErrorCode`), un `limit` opcional (`ReaderLimitName`) y `cause`. |
| `isExcelBridgeError` | `(error) => error is ExcelBridgeError`. Prefiérelo a `instanceof`. |

## Funciones auxiliares

| Export | Firma |
| --- | --- |
| `coordinateToIndex` | `(coordinate) => { row, col }`, así que `"A1"` es `{ row: 0, col: 0 }`. |
| `indexToCoordinate` | `(row, col) => string`, así que `(0, 0)` es `"A1"`. |
| `dateToExcelSerial` | `(date) => number` |
| `excelSerialToDate` | `(serial) => Date` |
| `isDate` | `(value) => value is Date`, verdadero solo para un `Date` válido. |
| `calculateColumnWidths` | `(data) => number[]`, los anchos que usaría `autoWidth`. |
| `EXCEL_LIMITS` | Los límites propios de Excel: filas, columnas, longitud de celda, hipervínculos. |

## ExcelBridge

`ExcelBridge` reúne `read`, `readFromFile`, `write`, `writeBuffer`, `Reader`, `Writer`, `Workbook`, `coordinateToIndex` y `indexToCoordinate` en un solo objeto. Sirve para scripts y prototipos.

> [!NOTE]
> Los empaquetadores conservan un objeto completo, así que incluso una sola llamada a `ExcelBridge.write` incluye también el lector. En el código del navegador, importa las exportaciones con nombre.

## Tipos

Los tipos que más usarás. Todos los tipos se exportan desde la raíz del paquete.

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

Las otras formas de celda, `ParsedStringCell`, `ParsedBooleanCell`, `ParsedDateCell` y `ParsedEmptyCell`, siguen la tabla de [Lee un libro](../guide/reading/). `ParsedRow` es un arreglo de `ParsedCell`. `ParsedBorder` es `true` o `BorderLines`, un objeto con una entrada `BorderLine` por cada lado que tiene línea.

### Todos los demás tipos

`AutoFilter`, `BorderLines`, `BorderSideName`, `CellValidation`, `CellValueConditionalFormat`, `ColorScaleConditionalFormat`, `ConditionalFormat`, `ConditionalFormatOperator`, `ConditionalFormatStyle`, `DataValidationOperator`, `DataValidationType`, `ExcelBridgeErrorCode`, `ExcelBridgeErrorOptions`, `ExcelReaderOptions`, `ExcelWriterOptions`, `ExpressionConditionalFormat`, `ExternalHyperlink`, `Hyperlink`, `HyperlinkOptions`, `InternalHyperlink`, `ObjectCellValue`, `ObjectsToSheetOptions`, `ObjectsToStreamingSheetOptions`, `ReadColumn`, `ReaderLimitName`, `SheetState`, `SheetToObjectsOptions`, `StreamingColumn`, `WorkbookMetadata` y `WriteColumn`.
