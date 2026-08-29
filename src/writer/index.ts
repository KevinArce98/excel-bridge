import { createExcelBlob, createExcelBuffer, ExcelFiles } from '../core/zip-manager';
import {
  generateSheetXml,
  generateStylesXml,
  generateSharedStringsXml,
  generateContentTypesXml,
  generateWorkbookXml,
  generateWorkbookRelsXml,
  generateRootRelsXml,
  generateCorePropsXml,
  generateAppPropsXml,
  SheetGenerationOptions,
} from '../core/xml-templates';
import { StyleManager } from '../core/style-manager';
import { isDate } from '../core/date-utils';
import { CellValue, CellValidation, CellStyle, ConditionalFormat } from '../core/types';

export type {
  CellValue,
  CellValidation,
  CellStyle,
  ConditionalFormat,
  DataValidationType,
  DataValidationOperator,
} from '../core/types';
export { dataValidation } from './validation';

export interface SheetOptions {
  name?: string;
  freezePane?: { row?: number; col?: number };
  autoWidth?: boolean;
  columnWidths?: number[];
}

export interface ExcelWriterOptions {
  creator?: string;
  title?: string;
  subject?: string;
  sharedStrings?: boolean;
}

export interface ExcelData {
  data: CellValue[][];
  validations?: CellValidation[];
  styles?: Record<string, CellStyle>;
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  options?: SheetOptions;
}

export class ExcelWriter {
  private _options: ExcelWriterOptions;

  constructor(options: ExcelWriterOptions = {}) {
    this._options = {
      creator: 'Excel Bridge',
      ...options,
    };
  }

  createWorkbook(data: ExcelData[]): Blob {
    const files = this.generateFiles(data);
    return createExcelBlob(files);
  }

  createWorkbookBuffer(data: ExcelData[]): Uint8Array {
    const files = this.generateFiles(data);
    return createExcelBuffer(files);
  }

  private generateFiles(data: ExcelData[]): ExcelFiles {
    const sheetCount = data.length;

    const shared = this._options.sharedStrings ? this.buildSharedStrings(data) : null;
    const hasSharedStrings = !!shared && shared.list.length > 0;

    const styleManager = new StyleManager();

    data.forEach(sheetData => {
      if (sheetData.styles) {
        Object.values(sheetData.styles).forEach(style => {
          styleManager.getStyleId(style);
        });
      }
    });

    const containsDates = data.some(sheetData =>
      sheetData.data.some(row => row.some(cell => isDate(cell)))
    );

    if (containsDates) {
      styleManager.getDateStyleId();
    }

    const sheetNames = data.map((sheet, index) => sheet.options?.name || `Sheet${index + 1}`);

    const worksheetEntries: Array<{ path: string; xml: string }> = [];

    data.forEach((sheetData, index) => {
      const sheetIndex = index + 1;
      const sheetOptions: SheetGenerationOptions = {
        freezePane: sheetData.options?.freezePane,
        autoWidth: sheetData.options?.autoWidth,
        columnWidths: sheetData.options?.columnWidths,
        mergeCells: sheetData.mergeCells,
        conditionalFormats: sheetData.conditionalFormats,
        sharedStrings: hasSharedStrings ? shared!.map : undefined,
      };

      const sheetXml = generateSheetXml(
        sheetData.data,
        sheetData.validations || [],
        sheetData.styles || {},
        styleManager,
        sheetOptions
      );

      worksheetEntries.push({ path: `xl/worksheets/sheet${sheetIndex}.xml`, xml: sheetXml });
    });

    const files: ExcelFiles = {};

    files['[Content_Types].xml'] = generateContentTypesXml(sheetCount, hasSharedStrings);

    files['_rels/.rels'] = generateRootRelsXml();

    files['xl/_rels/workbook.xml.rels'] = generateWorkbookRelsXml(sheetCount, hasSharedStrings);

    files['xl/workbook.xml'] = generateWorkbookXml(sheetNames);

    files['xl/styles.xml'] = generateStylesXml(styleManager);

    worksheetEntries.forEach(entry => {
      files[entry.path] = entry.xml;
    });

    if (hasSharedStrings) {
      files['xl/sharedStrings.xml'] = generateSharedStringsXml(shared!.list);
    }

    files['docProps/core.xml'] = generateCorePropsXml(
      this._options.creator,
      this._options.title,
      this._options.subject
    );
    files['docProps/app.xml'] = generateAppPropsXml();

    return files;
  }

  private buildSharedStrings(data: ExcelData[]): { map: Map<string, number>; list: string[] } {
    const map = new Map<string, number>();
    const list: string[] = [];

    data.forEach(sheetData => {
      sheetData.data.forEach(row => {
        row.forEach(cell => {
          if (typeof cell === 'string' && !cell.startsWith('=') && !map.has(cell)) {
            map.set(cell, list.length);
            list.push(cell);
          }
        });
      });
    });

    return { map, list };
  }

  addValidation(data: ExcelData[], range: string, options: string): ExcelData[] {
    const newData = [...data];
    const lastSheet = newData[newData.length - 1];

    if (lastSheet) {
      if (!lastSheet.validations) {
        lastSheet.validations = [];
      }
      lastSheet.validations.push({ range, options });
    }

    return newData;
  }

  addStyle(data: ExcelData[], rowIndex: number, colIndex: number, style: CellStyle): ExcelData[] {
    const newData = [...data];
    const lastSheet = newData[newData.length - 1];

    if (lastSheet) {
      if (!lastSheet.styles) {
        lastSheet.styles = {};
      }
      lastSheet.styles[`${rowIndex}-${colIndex}`] = style;
    }

    return newData;
  }

  static createSimple(data: CellValue[][], options?: ExcelWriterOptions): Blob {
    const writer = new ExcelWriter(options);
    return writer.createWorkbook([{ data }]);
  }

  static createSimpleBuffer(data: CellValue[][], options?: ExcelWriterOptions): Uint8Array {
    const writer = new ExcelWriter(options);
    return writer.createWorkbookBuffer([{ data }]);
  }
}

export const createExcelFile = (data: CellValue[][], options?: ExcelWriterOptions): Blob => {
  return ExcelWriter.createSimple(data, options);
};

export const createExcelFileBuffer = (
  data: CellValue[][],
  options?: ExcelWriterOptions
): Uint8Array => {
  return ExcelWriter.createSimpleBuffer(data, options);
};
