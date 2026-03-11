import { XMLParser } from 'fast-xml-parser';
import { extractExcelFiles, validateExcelStructure } from '../core/zip-manager';

export interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
}

export interface ParsedSheet {
  name: string;
  data: ParsedCell[][];
  validations: Array<{
    range: string;
    options: string;
  }>;
}

export interface ParsedWorkbook {
  sheets: ParsedSheet[];
  metadata: {
    created?: string;
    modified?: string;
    creator?: string;
  };
}

export class ExcelReader {
  private parser: XMLParser;

  constructor() {
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
      textNodeName: '#text',
      parseAttributeValue: true,
      parseTagValue: true,
    });
  }

  async parseFromFile(file: File): Promise<ParsedWorkbook> {
    const buffer = new Uint8Array(await file.arrayBuffer());
    return this.parseFromBuffer(buffer);
  }

  parseFromBuffer(buffer: Uint8Array): ParsedWorkbook {
    try {
      const files = extractExcelFiles(buffer);

      if (!validateExcelStructure(files)) {
        throw new Error('Invalid Excel file structure');
      }

      const workbookXml = files['xl/workbook.xml'];
      const workbook = this.parser.parse(workbookXml);

      const sharedStrings = this.parseSharedStrings(files);

      const sheets: ParsedSheet[] = [];
      const sheetElements = workbook.workbook?.sheets?.sheet || [];

      if (Array.isArray(sheetElements)) {
        for (const sheetElement of sheetElements) {
          const sheetName = sheetElement.name || 'Sheet1';
          const sheetPath = `xl/worksheets/sheet${sheetElement.sheetId || 1}.xml`;

          if (files[sheetPath]) {
            const sheetData = this.parseSheet(files[sheetPath], sharedStrings);
            sheets.push({
              name: sheetName,
              ...sheetData,
            });
          }
        }
      } else if (sheetElements) {
        const sheetName = sheetElements.name || 'Sheet1';
        const sheetData = this.parseSheet(files['xl/worksheets/sheet1.xml'], sharedStrings);
        sheets.push({
          name: sheetName,
          ...sheetData,
        });
      }

      return {
        sheets,
        metadata: this.extractMetadata(files),
      };
    } catch (error) {
      throw new Error(
        `Failed to parse Excel file: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  private parseSharedStrings(files: Record<string, string>): string[] {
    const sharedStringsXml = files['xl/sharedStrings.xml'];
    if (!sharedStringsXml) {
      return [];
    }

    try {
      const parsed = this.parser.parse(sharedStringsXml);
      const sst = parsed.sst;

      if (!sst || !sst.si) {
        return [];
      }

      const items = Array.isArray(sst.si) ? sst.si : [sst.si];
      return items.map((item: any) => {
        if (item.t) {
          return item.t['#text'] || item.t || '';
        }
        return '';
      });
    } catch {
      return [];
    }
  }

  private parseSheet(sheetXml: string, sharedStrings: string[]): Omit<ParsedSheet, 'name'> {
    const parsed = this.parser.parse(sheetXml);
    const worksheet = parsed.worksheet;

    const rows = worksheet?.sheetData?.row || [];
    const validations = worksheet?.dataValidations?.dataValidation || [];

    const data: ParsedCell[][] = [];
    const parsedValidations = [];

    if (Array.isArray(validations)) {
      for (const validation of validations) {
        parsedValidations.push({
          range: validation.sqref,
          options: validation.formula1?.replace(/"/g, '') || '',
        });
      }
    } else if (validations) {
      parsedValidations.push({
        range: validations.sqref,
        options: validations.formula1?.replace(/"/g, '') || '',
      });
    }

    if (Array.isArray(rows)) {
      for (const rowElement of rows) {
        const rowIndex = parseInt(rowElement.r) - 1;
        const cells = rowElement.c || [];
        const rowData: ParsedCell[] = [];

        if (Array.isArray(cells)) {
          for (const cell of cells) {
            const cellData = this.parseCell(cell, rowIndex, sharedStrings);
            rowData.push(cellData);
          }
        } else if (cells) {
          const cellData = this.parseCell(cells, rowIndex, sharedStrings);
          rowData.push(cellData);
        }

        data.push(rowData);
      }
    } else if (rows) {
      const rowIndex = parseInt(rows.r) - 1;
      const cells = rows.c || [];
      const rowData: ParsedCell[] = [];

      if (Array.isArray(cells)) {
        for (const cell of cells) {
          const cellData = this.parseCell(cell, rowIndex, sharedStrings);
          rowData.push(cellData);
        }
      } else if (cells) {
        const cellData = this.parseCell(cells, rowIndex, sharedStrings);
        rowData.push(cellData);
      }

      data.push(rowData);
    }

    return {
      data,
      validations: parsedValidations,
    };
  }

  private parseCell(cell: any, rowIndex: number, sharedStrings: string[]): ParsedCell {
    const coordinate = cell.r;
    const columnIndex = this.columnLetterToIndex(coordinate.replace(/\d+/, ''));

    let value: any = null;
    let type: ParsedCell['type'] = 'empty';

    if (cell.v !== undefined) {
      value = cell.v;

      if (cell.t === 'b') {
        value = value === '1' || value === 1 || value === true;
        type = 'boolean';
      } else if (cell.t === 'inlineStr') {
        value = cell.is?.t?.['#text'] || cell.is?.t || '';
        type = 'string';
      } else if (cell.t === 's') {
        const stringIndex = parseInt(cell.v);
        value = sharedStrings[stringIndex] || '';
        type = 'string';
      } else if (cell.t === 'str') {
        value = value.toString();
        type = 'string';
      } else {
        value = parseFloat(value);
        type = 'number';
      }
    } else if (cell.is?.t) {
      value = cell.is.t['#text'] || cell.is.t || '';
      type = 'string';
    }

    return {
      value,
      type,
      coordinate,
      rowIndex,
      columnIndex,
    };
  }

  private columnLetterToIndex(letters: string): number {
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  private extractMetadata(files: Record<string, string>): ParsedWorkbook['metadata'] {
    const metadata: ParsedWorkbook['metadata'] = {};

    try {
      const appXml = files['docProps/app.xml'];
      if (appXml) {
        const parsed = this.parser.parse(appXml);
        const properties = parsed.Properties;

        if (properties) {
          metadata.creator = properties.Creator;
          metadata.created = properties.Created;
          metadata.modified = properties.Modified;
        }
      }
    } catch {
      // Ignore metadata parsing errors
    }

    return metadata;
  }
}

export const parseExcel = (buffer: Uint8Array): ParsedWorkbook => {
  const reader = new ExcelReader();
  return reader.parseFromBuffer(buffer);
};
