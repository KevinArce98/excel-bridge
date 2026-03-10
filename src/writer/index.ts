import { createExcelBlob, createExcelBuffer, ExcelFiles } from '../core/zip-manager';
import {
  generateSheetXml,
  generateSharedStringsXml,
  generateStylesXml,
  generateContentTypesXml,
  generateWorkbookXml,
  generateWorkbookRelsXml,
  generateRootRelsXml,
} from '../core/xml-templates';

export interface CellValidation {
  range: string;
  options: string;
}

export interface CellStyle {
  background?: string;
  border?: boolean;
  bold?: boolean;
  color?: string;
}

export interface ExcelWriterOptions {
  sheetName?: string;
  creator?: string;
  title?: string;
  subject?: string;
}

export interface ExcelData {
  data: any[][];
  validations?: CellValidation[];
  styles?: Record<string, CellStyle>;
}

export class ExcelWriter {
  private _options: ExcelWriterOptions;

  constructor(options: ExcelWriterOptions = {}) {
    this._options = {
      sheetName: 'Sheet1',
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
    const files: ExcelFiles = {};

    // Content Types
    files['[Content_Types].xml'] = generateContentTypesXml();

    // Root relationships
    files['_rels/.rels'] = generateRootRelsXml();

    // Workbook
    files['xl/workbook.xml'] = generateWorkbookXml();
    files['xl/_rels/workbook.xml.rels'] = generateWorkbookRelsXml();

    // Styles (shared across all sheets)
    files['xl/styles.xml'] = generateStylesXml();

    // Generate shared strings from all data
    const allStrings = this.extractAllStrings(data);
    if (allStrings.length > 0) {
      files['xl/sharedStrings.xml'] = generateSharedStringsXml(allStrings);
    }

    // Generate worksheets
    data.forEach((sheetData, index) => {
      const sheetIndex = index + 1;
      files[`xl/worksheets/sheet${sheetIndex}.xml`] = generateSheetXml(
        sheetData.data,
        sheetData.validations || [],
        sheetData.styles || {}
      );
    });

    return files;
  }

  private extractAllStrings(data: ExcelData[]): string[] {
    const strings: string[] = [];

    data.forEach(sheetData => {
      sheetData.data.forEach(row => {
        row.forEach(cell => {
          if (typeof cell === 'string') {
            strings.push(cell);
          } else if (cell !== null && cell !== undefined && typeof cell.toString === 'function') {
            const str = cell.toString();
            if (str && str !== '' && !isNaN(Date.parse(str))) {
              // Don't treat date strings as shared strings for now
              return;
            }
            if (str && str !== '') {
              strings.push(str);
            }
          }
        });
      });
    });

    return strings;
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

  static createSimple(data: any[][], options?: ExcelWriterOptions): Blob {
    const writer = new ExcelWriter(options);
    return writer.createWorkbook([{ data }]);
  }

  static createSimpleBuffer(data: any[][], options?: ExcelWriterOptions): Uint8Array {
    const writer = new ExcelWriter(options);
    return writer.createWorkbookBuffer([{ data }]);
  }
}

export const createExcelFile = (data: any[][], options?: ExcelWriterOptions): Blob => {
  return ExcelWriter.createSimple(data, options);
};

export const createExcelFileBuffer = (data: any[][], options?: ExcelWriterOptions): Uint8Array => {
  return ExcelWriter.createSimpleBuffer(data, options);
};
