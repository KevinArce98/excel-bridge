import { XMLParser } from 'fast-xml-parser';
import { extractExcelFiles, validateExcelStructure } from '../core/zip-manager';
import { excelSerialToDate, isDateNumFmtId } from '../core/date-utils';

export interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'date' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
  /** Present when the cell holds a formula (without the leading '='). */
  formula?: string;
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

type DateStyleIndices = Set<number>;

const toArray = <T>(value: T | T[] | undefined): T[] => {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
};

export class ExcelReader {
  private parser: XMLParser;

  constructor() {
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
      textNodeName: '#text',
      parseAttributeValue: true,
      parseTagValue: true,
      trimValues: false, // preserve intentional leading/trailing whitespace in strings
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

      const workbook = this.parser.parse(files['xl/workbook.xml']);
      const sharedStrings = this.parseSharedStrings(files);
      const dateStyles = this.parseDateStyles(files);
      const relMap = this.parseWorkbookRels(files);

      const sheets: ParsedSheet[] = [];
      const sheetElements = toArray(workbook.workbook?.sheets?.sheet);

      sheetElements.forEach((sheetElement, index) => {
        const sheetName = sheetElement.name ?? `Sheet${index + 1}`;
        const sheetPath = this.resolveSheetPath(sheetElement, relMap, index);

        if (sheetPath && files[sheetPath]) {
          const sheetData = this.parseSheet(files[sheetPath], sharedStrings, dateStyles);
          sheets.push({ name: String(sheetName), ...sheetData });
        }
      });

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

  /** Map relationship ids (r:id) to their target part paths, resolved relative to xl/. */
  private parseWorkbookRels(files: Record<string, string>): Record<string, string> {
    const relsXml = files['xl/_rels/workbook.xml.rels'];
    const map: Record<string, string> = {};
    if (!relsXml) return map;

    try {
      const parsed = this.parser.parse(relsXml);
      const rels = toArray(parsed.Relationships?.Relationship);
      for (const rel of rels) {
        if (!rel.Id || !rel.Target) continue;
        let target: string = String(rel.Target);
        if (target.startsWith('/')) {
          target = target.slice(1); // absolute package path
        } else {
          target = `xl/${target}`; // relative to the workbook part
        }
        map[String(rel.Id)] = target;
      }
    } catch {
      // Ignore malformed rels; fall back to sheetId-based resolution.
    }

    return map;
  }

  /**
   * Resolve a worksheet's part path. Prefer the relationship target (r:id),
   * falling back to sheetId, then positional ordering.
   */
  private resolveSheetPath(
    sheetElement: any,
    relMap: Record<string, string>,
    index: number
  ): string | undefined {
    const rId = sheetElement['r:id'] ?? sheetElement.id;
    if (rId && relMap[String(rId)]) {
      return relMap[String(rId)];
    }

    const sheetId = sheetElement.sheetId ?? index + 1;
    return `xl/worksheets/sheet${sheetId}.xml`;
  }

  private parseSharedStrings(files: Record<string, string>): string[] {
    const sharedStringsXml = files['xl/sharedStrings.xml'];
    if (!sharedStringsXml) {
      return [];
    }

    try {
      const parsed = this.parser.parse(sharedStringsXml);
      const items = toArray(parsed.sst?.si);
      return items.map((item: any) => this.extractStringItem(item));
    } catch {
      return [];
    }
  }

  /** Extract text from a shared-string <si>, handling rich-text runs (<r>). */
  private extractStringItem(item: any): string {
    if (item == null) return '';
    if (typeof item === 'string' || typeof item === 'number') return String(item);

    if (item.t !== undefined) {
      return this.extractText(item.t);
    }
    // Rich text: concatenate the text of each run.
    if (item.r !== undefined) {
      return toArray(item.r)
        .map((run: any) => this.extractText(run?.t))
        .join('');
    }
    return '';
  }

  private extractText(t: any): string {
    if (t == null) return '';
    if (typeof t === 'object') {
      return t['#text'] !== undefined ? String(t['#text']) : '';
    }
    return String(t);
  }

  /**
   * Parse styles.xml to find which cell-style indices map to a date/time number format.
   */
  private parseDateStyles(files: Record<string, string>): DateStyleIndices {
    const dateStyles: DateStyleIndices = new Set();
    const stylesXml = files['xl/styles.xml'];
    if (!stylesXml) return dateStyles;

    try {
      const parsed = this.parser.parse(stylesXml);
      const styleSheet = parsed.styleSheet;
      if (!styleSheet) return dateStyles;

      const customFormats: Record<number, string> = {};
      for (const fmt of toArray(styleSheet.numFmts?.numFmt)) {
        if (fmt.numFmtId !== undefined && fmt.formatCode !== undefined) {
          customFormats[Number(fmt.numFmtId)] = String(fmt.formatCode);
        }
      }

      const xfs = toArray(styleSheet.cellXfs?.xf);
      xfs.forEach((xf: any, index: number) => {
        const numFmtId = xf?.numFmtId !== undefined ? Number(xf.numFmtId) : 0;
        if (isDateNumFmtId(numFmtId, customFormats)) {
          dateStyles.add(index);
        }
      });
    } catch {
      // Ignore malformed styles.
    }

    return dateStyles;
  }

  private parseSheet(
    sheetXml: string,
    sharedStrings: string[],
    dateStyles: DateStyleIndices
  ): Omit<ParsedSheet, 'name'> {
    const parsed = this.parser.parse(sheetXml);
    const worksheet = parsed.worksheet;

    const rows = toArray(worksheet?.sheetData?.row);
    const validations = toArray(worksheet?.dataValidations?.dataValidation);

    const parsedValidations = validations.map((validation: any) => ({
      range: validation.sqref,
      options: this.extractText(validation.formula1).replace(/"/g, '') || '',
    }));

    const data: ParsedCell[][] = [];

    for (const rowElement of rows) {
      const rowIndex = parseInt(rowElement.r, 10) - 1;
      const cells = toArray(rowElement.c);

      // Place each cell at its real column index so sparse rows stay aligned.
      const rowData: ParsedCell[] = [];
      let maxCol = -1;

      for (const cell of cells) {
        const parsedCell = this.parseCell(cell, rowIndex, sharedStrings, dateStyles);
        rowData[parsedCell.columnIndex] = parsedCell;
        maxCol = Math.max(maxCol, parsedCell.columnIndex);
      }

      // Fill any gaps with explicit empty cells.
      for (let c = 0; c <= maxCol; c++) {
        if (!rowData[c]) {
          rowData[c] = {
            value: null,
            type: 'empty',
            coordinate: `${this.columnIndexToLetter(c)}${rowIndex + 1}`,
            rowIndex,
            columnIndex: c,
          };
        }
      }

      data.push(rowData);
    }

    return {
      data,
      validations: parsedValidations,
    };
  }

  private parseCell(
    cell: any,
    rowIndex: number,
    sharedStrings: string[],
    dateStyles: DateStyleIndices
  ): ParsedCell {
    const coordinate = String(cell.r ?? '');
    const columnIndex = this.columnLetterToIndex(coordinate.replace(/\d+/g, ''));
    const styleIndex = cell.s !== undefined ? Number(cell.s) : undefined;

    let value: any = null;
    let type: ParsedCell['type'] = 'empty';
    let formula: string | undefined;

    if (cell.f !== undefined) {
      formula = this.extractText(cell.f);
    }

    if (cell.t === 'inlineStr' || (cell.t === undefined && cell.is !== undefined)) {
      value = this.extractText(cell.is?.t);
      type = 'string';
    } else if (cell.v !== undefined) {
      const raw = cell.v;

      if (cell.t === 'b') {
        value = raw === '1' || raw === 1 || raw === true;
        type = 'boolean';
      } else if (cell.t === 's') {
        value = sharedStrings[parseInt(raw, 10)] ?? '';
        type = 'string';
      } else if (cell.t === 'str') {
        value = String(raw);
        type = 'string';
      } else {
        // Numeric cell. Convert to a Date when the style says so.
        const num = typeof raw === 'number' ? raw : parseFloat(raw);
        if (styleIndex !== undefined && dateStyles.has(styleIndex) && !isNaN(num)) {
          value = excelSerialToDate(num);
          type = 'date';
        } else {
          value = num;
          type = 'number';
        }
      }
    }

    return {
      value,
      type,
      coordinate,
      rowIndex,
      columnIndex,
      ...(formula !== undefined ? { formula } : {}),
    };
  }

  private columnLetterToIndex(letters: string): number {
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  private columnIndexToLetter(index: number): string {
    let letter = '';
    let num = index + 1;
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      num = Math.floor((num - 1) / 26);
    }
    return letter;
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

      // Core properties carry the canonical created/modified timestamps and creator.
      const coreXml = files['docProps/core.xml'];
      if (coreXml) {
        const parsed = this.parser.parse(coreXml);
        const core = parsed['cp:coreProperties'];
        if (core) {
          metadata.creator = this.extractText(core['dc:creator']) || metadata.creator;
          metadata.created = this.extractText(core['dcterms:created']) || metadata.created;
          metadata.modified = this.extractText(core['dcterms:modified']) || metadata.modified;
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
