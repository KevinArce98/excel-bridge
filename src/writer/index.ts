import { createExcelBlob, createExcelBuffer, ExcelFiles } from '../core/zip-manager';
import {
  generateSheetXml,
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
    // CRITICAL: Force hasSharedStrings to false
    // We use inlineStr in cells, so we must NOT declare sharedStrings.xml
    // Having sharedStrings.xml declared but using inlineStr causes Excel to reject the file
    const hasSharedStrings = false;

    // CRITICAL: Order matters for Excel compatibility, especially on Mac
    // [Content_Types].xml MUST be first in the ZIP index
    const files: ExcelFiles = {};

    // 1. Content Types - MUST BE FIRST
    files['[Content_Types].xml'] = generateContentTypesXml(hasSharedStrings);

    // 2. Root relationships
    files['_rels/.rels'] = generateRootRelsXml();

    // 3. Workbook relationships
    files['xl/_rels/workbook.xml.rels'] = generateWorkbookRelsXml(hasSharedStrings);

    // 4. Workbook
    files['xl/workbook.xml'] = generateWorkbookXml();

    // 5. Styles
    files['xl/styles.xml'] = generateStylesXml();

    // 6. NO Shared strings - we use inlineStr instead

    // 7. Worksheets
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
