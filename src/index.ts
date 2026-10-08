export { ExcelReader, parseExcel } from './reader';
export type {
  ExcelReaderOptions,
  ParsedCell,
  ParsedStringCell,
  ParsedNumberCell,
  ParsedBooleanCell,
  ParsedDateCell,
  ParsedErrorCell,
  ParsedEmptyCell,
  ParsedRow,
  ParsedSheet,
  ParsedWorkbook,
} from './reader';
export { Workbook } from './workbook';
export type { WorkbookMetadata } from './workbook';

export {
  ExcelWriter,
  createExcelFile,
  createExcelFileBuffer,
  dataValidation,
  hyperlink,
} from './writer';
export type {
  ExcelWriterOptions,
  ExcelData,
  CellValidation,
  CellStyle,
  CellValue,
  FormulaCell,
  FormulaResult,
  TextCell,
  ErrorCell,
  ExcelErrorValue,
  SheetOptions,
  ConditionalFormat,
  DataValidationType,
  DataValidationOperator,
  AutoFilter,
  Hyperlink,
  HyperlinkOptions,
  SheetLayout,
  SheetState,
} from './writer';
export type {
  CellBorder,
  BorderSide,
  BorderSides,
  BorderLine,
  BorderLines,
  BorderSideName,
  BorderStyleName,
  ParsedBorder,
  ConditionalFormatStyle,
  ConditionalFormatOperator,
  CellValueConditionalFormat,
  ExpressionConditionalFormat,
  ColorScaleConditionalFormat,
  ExternalHyperlink,
  InternalHyperlink,
} from './core/types';

export { createExcelWorkbookStream, streamToBuffer } from './writer/stream';
export type { StreamingSheetInput } from './writer/stream';

export { dateToExcelSerial, excelSerialToDate, isDate, EXCEL_LIMITS } from './core/date-utils';

export { isExcelErrorValue } from './core/cells';

export { calculateColumnWidths } from './core/column-width';

import { coordinateToIndex, indexToCoordinate } from './core/cell-ref';
export { coordinateToIndex, indexToCoordinate };

import { ExcelReader as ReaderClass, parseExcel as parseFunction } from './reader';
import type { ExcelReaderOptions } from './reader';
import {
  ExcelWriter as WriterClass,
  createExcelFile as createFile,
  createExcelFileBuffer as createFileBuffer,
} from './writer';
import { Workbook as WorkbookClass } from './workbook';

export const ExcelBridge = {
  read: parseFunction,
  readFromFile: (file: File, options?: ExcelReaderOptions) => {
    const reader = new ReaderClass(options);
    return reader.parseFromFile(file);
  },

  write: createFile,
  writeBuffer: createFileBuffer,

  coordinateToIndex,
  indexToCoordinate,

  Writer: WriterClass,
  Reader: ReaderClass,
  Workbook: WorkbookClass,
};
export { ExcelBridgeError, isExcelBridgeError } from './core/errors';
export type { ExcelBridgeErrorCode, ExcelBridgeErrorOptions, ReaderLimitName } from './core/errors';

export { sheetToObjects } from './objects/read';
export type { ObjectCellValue, ReadColumn, SheetToObjectsOptions } from './objects/read';
export { objectsToSheet, objectsToStreamingSheet } from './objects/write';
export type {
  ObjectsToSheetOptions,
  ObjectsToStreamingSheetOptions,
  StreamingColumn,
  WriteColumn,
} from './objects/write';

export { XLSX_CONTENT_TYPE } from './core/constants';
export { downloadXlsx } from './delivery/download';
export { toReadableStream, xlsxResponse } from './delivery/response';
