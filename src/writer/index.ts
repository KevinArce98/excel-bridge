import { createExcelBlob, createExcelBuffer, ExcelFiles } from '../core/zip-manager';
import {
  generateSheetXml,
  generateStylesXml,
  generateWorkbookRelsXml,
  SheetGenerationOptions,
} from '../core/xml-templates';
import { StyleManager } from '../core/style-manager';
import { isDate } from '../core/date-utils';

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

export interface SheetOptions {
  name?: string;
  freezePane?: { row?: number; col?: number };
  autoWidth?: boolean;
}

export interface ExcelWriterOptions {
  creator?: string;
  title?: string;
  subject?: string;
}

export interface ExcelData {
  data: any[][];
  validations?: CellValidation[];
  styles?: Record<string, CellStyle>;
  mergeCells?: string[];
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
    const hasSharedStrings = false;
    const sheetCount = data.length;

    // Create a single StyleManager for all sheets
    const styleManager = new StyleManager();

    // Pre-process all styles to populate StyleManager
    data.forEach(sheetData => {
      if (sheetData.styles) {
        Object.values(sheetData.styles).forEach(style => {
          styleManager.getStyleId(style);
        });
      }
    });

    // Ensure date style is registered before generating styles.xml
    const containsDates = data.some(sheetData =>
      sheetData.data.some(row => row.some(cell => isDate(cell)))
    );

    if (containsDates) {
      styleManager.getDateStyleId();
    }

    // Extract sheet names
    const sheetNames = data.map((sheet, index) => sheet.options?.name || `Sheet${index + 1}`);

    // Generate worksheet XML first to capture style usage
    const worksheetEntries: Array<{ path: string; xml: string }> = [];

    data.forEach((sheetData, index) => {
      const sheetIndex = index + 1;
      const sheetOptions: SheetGenerationOptions = {
        freezePane: sheetData.options?.freezePane,
        autoWidth: sheetData.options?.autoWidth,
        mergeCells: sheetData.mergeCells,
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

    // SAFE FUNCTIONAL VERSION - NO DOCPROPS
    const files: ExcelFiles = {};

    // 1. [Content_Types].xml - MUST BE FIRST
    files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${Array.from({ length: sheetCount }, (_, i) => `  <Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

    // 2. _rels/.rels
    files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    // 3. xl/_rels/workbook.xml.rels
    files['xl/_rels/workbook.xml.rels'] = generateWorkbookRelsXml(sheetCount, hasSharedStrings);

    // 4. xl/workbook.xml - SIMPLE VERSION
    files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${sheetNames.map((name, index) => `    <sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('\n')}
  </sheets>
</workbook>`;

    // 5. xl/styles.xml (after worksheets so StyleManager has all styles)
    files['xl/styles.xml'] = generateStylesXml(styleManager);

    // 6. xl/worksheets/sheet*.xml
    worksheetEntries.forEach(entry => {
      files[entry.path] = entry.xml;
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
