// Reader exports
export { ExcelReader, parseExcel } from './reader';
export type { ParsedCell, ParsedSheet, ParsedWorkbook } from './reader';

// Writer exports
export { ExcelWriter, createExcelFile, createExcelFileBuffer } from './writer';
export type { ExcelWriterOptions, ExcelData, CellValidation, CellStyle } from './writer';

// Core utilities
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
} from './core/xml-templates';

// Utility functions
export const coordinateToIndex = (coordinate: string): { row: number; col: number } => {
  const match = coordinate.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid coordinate format: ${coordinate}`);
  }

  const colLetters = match[1];
  const rowNumber = parseInt(match[2]) - 1;

  let colIndex = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 64);
  }
  colIndex -= 1;

  return { row: rowNumber, col: colIndex };
};

export const indexToCoordinate = (row: number, col: number): string => {
  const rowNumber = row + 1;
  let colLetters = '';

  let colIndex = col + 1;
  while (colIndex > 0) {
    const remainder = (colIndex - 1) % 26;
    colLetters = String.fromCharCode(65 + remainder) + colLetters;
    colIndex = Math.floor((colIndex - 1) / 26);
  }

  return `${colLetters}${rowNumber}`;
};

// Main API for quick usage
import { ExcelReader as ReaderClass, parseExcel as parseFunction } from './reader';
import {
  ExcelWriter as WriterClass,
  createExcelFile as createFile,
  createExcelFileBuffer as createFileBuffer,
} from './writer';

export const ExcelBridge = {
  // Reading
  read: parseFunction,
  readFromFile: (file: File) => {
    const reader = new ReaderClass();
    return reader.parseFromFile(file);
  },

  // Writing
  write: createFile,
  writeBuffer: createFileBuffer,

  // Utilities
  coordinateToIndex,
  indexToCoordinate,

  // Advanced
  Writer: WriterClass,
  Reader: ReaderClass,
};
