export { ExcelReader, parseExcel } from './reader';
export type { ParsedCell, ParsedSheet, ParsedWorkbook } from './reader';
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
  SheetOptions,
  ConditionalFormat,
  DataValidationType,
  DataValidationOperator,
  AutoFilter,
  Hyperlink,
  HyperlinkOptions,
} from './writer';
export type {
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

export {
  createExcelBlob,
  createExcelBuffer,
  extractExcelFiles,
  validateExcelStructure,
} from './core/zip-manager';
export type { ExcelFiles } from './core/zip-manager';

export { XML_NS, CONTENT_TYPES, RELATIONSHIP_TYPES, CELL_TYPES } from './core/constants';
export {
  generateSheetXml,
  generateSharedStringsXml,
  generateStylesXml,
  generateContentTypesXml,
  generateWorkbookXml,
  generateWorkbookRelsXml,
  generateRootRelsXml,
  generateCorePropsXml,
  generateAppPropsXml,
  generateSheetRelsXml,
} from './core/xml-templates';
export type { SheetGenerationOptions, DefinedName } from './core/xml-templates';

export { StyleManager } from './core/style-manager';
export type { ExcelStyle, Font, Fill, Border, CellAlignment } from './core/style-manager';

export {
  dateToExcelSerial,
  excelSerialToDate,
  isDate,
  isDateNumFmtId,
  isDateFormatCode,
  EXCEL_LIMITS,
  validateRowIndex,
  validateColIndex,
  validateCellValue,
} from './core/date-utils';

export { calculateColumnWidths, generateColsXml } from './core/column-width';

import { coordinateToIndex, indexToCoordinate } from './core/cell-ref';
export { coordinateToIndex, indexToCoordinate };

import { ExcelReader as ReaderClass, parseExcel as parseFunction } from './reader';
import {
  ExcelWriter as WriterClass,
  createExcelFile as createFile,
  createExcelFileBuffer as createFileBuffer,
} from './writer';
import { Workbook as WorkbookClass } from './workbook';

export const ExcelBridge = {
  read: parseFunction,
  readFromFile: (file: File) => {
    const reader = new ReaderClass();
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
