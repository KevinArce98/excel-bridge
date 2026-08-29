import { XMLParser } from 'fast-xml-parser';
import { extractExcelFiles, validateExcelStructure } from '../core/zip-manager';
import { excelSerialToDate, isDateNumFmtId } from '../core/date-utils';
import { CellStyle } from '../core/types';

export interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'date' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
  formula?: string;
}

export interface ParsedSheet {
  name: string;
  data: ParsedCell[][];
  validations: Array<{
    range: string;
    options: string;
  }>;
  styles?: Record<string, CellStyle>;
  mergeCells?: string[];
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
}

export interface ParsedWorkbook {
  sheets: ParsedSheet[];
  metadata: {
    created?: string;
    modified?: string;
    creator?: string;
    title?: string;
    subject?: string;
  };
}

interface DecodedFont {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  size?: number;
  name?: string;
}

interface DecodedFill {
  fgColor?: string;
  patternType?: string;
}

interface DecodedBorder {
  left?: boolean;
  right?: boolean;
  top?: boolean;
  bottom?: boolean;
}

interface DecodedXf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
}

interface StyleSheetData {
  fonts: DecodedFont[];
  fills: DecodedFill[];
  borders: DecodedBorder[];
  customFormats: Record<number, string>;
  cellXfs: DecodedXf[];
  dateStyles: Set<number>;
}

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
      trimValues: false,
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
      const styleSheet = this.parseStyleSheet(files);
      const relMap = this.parseWorkbookRels(files);

      const sheets: ParsedSheet[] = [];
      const sheetElements = toArray(workbook.workbook?.sheets?.sheet);

      sheetElements.forEach((sheetElement, index) => {
        const sheetName = sheetElement.name ?? `Sheet${index + 1}`;
        const sheetPath = this.resolveSheetPath(sheetElement, relMap, index);

        if (sheetPath && files[sheetPath]) {
          const sheetData = this.parseSheet(files[sheetPath], sharedStrings, styleSheet);
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
          target = target.slice(1);
        } else {
          target = `xl/${target}`;
        }
        map[String(rel.Id)] = target;
      }
    } catch {}

    return map;
  }

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

  private extractStringItem(item: any): string {
    if (item == null) return '';
    if (typeof item === 'string' || typeof item === 'number') return String(item);

    if (item.t !== undefined) {
      return this.extractText(item.t);
    }
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

  private parseStyleSheet(files: Record<string, string>): StyleSheetData {
    const result: StyleSheetData = {
      fonts: [],
      fills: [],
      borders: [],
      customFormats: {},
      cellXfs: [],
      dateStyles: new Set(),
    };

    const stylesXml = files['xl/styles.xml'];
    if (!stylesXml) return result;

    try {
      const parsed = this.parser.parse(stylesXml);
      const styleSheet = parsed.styleSheet;
      if (!styleSheet) return result;

      for (const fmt of toArray(styleSheet.numFmts?.numFmt)) {
        if (fmt.numFmtId !== undefined && fmt.formatCode !== undefined) {
          result.customFormats[Number(fmt.numFmtId)] = String(fmt.formatCode);
        }
      }

      result.fonts = toArray(styleSheet.fonts?.font).map((font: any) => ({
        bold: font?.b !== undefined,
        italic: font?.i !== undefined,
        underline: font?.u !== undefined,
        color: font?.color?.rgb !== undefined ? String(font.color.rgb) : undefined,
        size: font?.sz?.val !== undefined ? Number(font.sz.val) : undefined,
        name: font?.name?.val !== undefined ? String(font.name.val) : undefined,
      }));

      result.fills = toArray(styleSheet.fills?.fill).map((fill: any) => {
        const patternFill = fill?.patternFill;
        return {
          patternType: patternFill?.patternType,
          fgColor:
            patternFill?.fgColor?.rgb !== undefined ? String(patternFill.fgColor.rgb) : undefined,
        };
      });

      result.borders = toArray(styleSheet.borders?.border).map((border: any) => ({
        left: border?.left !== undefined && border.left.style !== undefined,
        right: border?.right !== undefined && border.right.style !== undefined,
        top: border?.top !== undefined && border.top.style !== undefined,
        bottom: border?.bottom !== undefined && border.bottom.style !== undefined,
      }));

      const xfs = toArray(styleSheet.cellXfs?.xf);
      result.cellXfs = xfs.map((xf: any) => {
        const alignment = xf?.alignment;
        return {
          fontId: xf?.fontId !== undefined ? Number(xf.fontId) : 0,
          fillId: xf?.fillId !== undefined ? Number(xf.fillId) : 0,
          borderId: xf?.borderId !== undefined ? Number(xf.borderId) : 0,
          numFmtId: xf?.numFmtId !== undefined ? Number(xf.numFmtId) : 0,
          alignment: alignment
            ? {
                horizontal: alignment.horizontal,
                vertical: alignment.vertical === 'center' ? 'middle' : alignment.vertical,
                wrapText: alignment.wrapText === 1 || alignment.wrapText === '1',
              }
            : undefined,
        };
      });

      result.cellXfs.forEach((xf, index) => {
        if (isDateNumFmtId(xf.numFmtId, result.customFormats)) {
          result.dateStyles.add(index);
        }
      });
    } catch {}

    return result;
  }

  private argbToHex(argb: string): string {
    const hex = argb.length === 8 ? argb.slice(2) : argb;
    return `#${hex}`;
  }

  private decodeCellStyle(
    styleIndex: number | undefined,
    styleSheet: StyleSheetData
  ): CellStyle | undefined {
    if (styleIndex === undefined || styleSheet.dateStyles.has(styleIndex)) {
      return undefined;
    }

    const xf = styleSheet.cellXfs[styleIndex];
    if (!xf) return undefined;

    const font = styleSheet.fonts[xf.fontId];
    const fill = styleSheet.fills[xf.fillId];
    const border = styleSheet.borders[xf.borderId];

    const style: CellStyle = {};

    if (font?.bold) style.bold = true;
    if (font?.italic) style.italic = true;
    if (font?.underline) style.underline = true;
    if (font?.color) style.color = this.argbToHex(font.color);
    if (font?.size) style.fontSize = font.size;
    if (font?.name) style.fontName = font.name;

    if (fill?.patternType === 'solid' && fill.fgColor) {
      style.background = this.argbToHex(fill.fgColor);
    }

    if (border && (border.left || border.right || border.top || border.bottom)) {
      style.border = true;
    }

    if (xf.alignment) {
      if (xf.alignment.horizontal) style.align = xf.alignment.horizontal;
      if (xf.alignment.vertical) style.verticalAlign = xf.alignment.vertical;
      if (xf.alignment.wrapText) style.wrapText = true;
    }

    if (xf.numFmtId) {
      const code = styleSheet.customFormats[xf.numFmtId];
      if (code) style.numberFormat = code;
    }

    return Object.keys(style).length > 0 ? style : undefined;
  }

  private parseSheet(
    sheetXml: string,
    sharedStrings: string[],
    styleSheet: StyleSheetData
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
    const styles: Record<string, CellStyle> = {};

    for (const rowElement of rows) {
      const rowIndex = parseInt(rowElement.r, 10) - 1;
      const cells = toArray(rowElement.c);

      const rowData: ParsedCell[] = [];
      let maxCol = -1;

      for (const cell of cells) {
        const parsedCell = this.parseCell(cell, rowIndex, sharedStrings, styleSheet);
        rowData[parsedCell.columnIndex] = parsedCell;
        maxCol = Math.max(maxCol, parsedCell.columnIndex);

        const styleIndex = cell.s !== undefined ? Number(cell.s) : undefined;
        const decodedStyle = this.decodeCellStyle(styleIndex, styleSheet);
        if (decodedStyle) {
          styles[`${rowIndex}-${parsedCell.columnIndex}`] = decodedStyle;
        }
      }

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

    const mergeCells = toArray(worksheet?.mergeCells?.mergeCell)
      .map((m: any) => (m?.ref !== undefined ? String(m.ref) : undefined))
      .filter((ref): ref is string => ref !== undefined);

    const freezePane = this.parseFreezePane(worksheet);
    const columnWidths = this.parseColumnWidths(worksheet);

    return {
      data,
      validations: parsedValidations,
      ...(Object.keys(styles).length > 0 ? { styles } : {}),
      ...(mergeCells.length > 0 ? { mergeCells } : {}),
      ...(freezePane ? { freezePane } : {}),
      ...(columnWidths ? { columnWidths } : {}),
    };
  }

  private parseFreezePane(worksheet: any): { row?: number; col?: number } | undefined {
    const sheetViews = toArray(worksheet?.sheetViews?.sheetView);
    const pane = sheetViews[0]?.pane;
    if (!pane) return undefined;

    const col = pane.xSplit !== undefined ? Number(pane.xSplit) : 0;
    const row = pane.ySplit !== undefined ? Number(pane.ySplit) : 0;
    if (!col && !row) return undefined;

    return { ...(row ? { row } : {}), ...(col ? { col } : {}) };
  }

  private parseColumnWidths(worksheet: any): number[] | undefined {
    const cols = toArray(worksheet?.cols?.col);
    if (cols.length === 0) return undefined;

    const widths: number[] = [];
    cols.forEach((col: any) => {
      const min = Number(col.min) - 1;
      const max = Number(col.max) - 1;
      const width = Number(col.width);
      for (let c = min; c <= max; c++) {
        widths[c] = width;
      }
    });

    return widths;
  }

  private parseCell(
    cell: any,
    rowIndex: number,
    sharedStrings: string[],
    styleSheet: StyleSheetData
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
        const num = typeof raw === 'number' ? raw : parseFloat(raw);
        if (styleIndex !== undefined && styleSheet.dateStyles.has(styleIndex) && !isNaN(num)) {
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

      const coreXml = files['docProps/core.xml'];
      if (coreXml) {
        const parsed = this.parser.parse(coreXml);
        const core = parsed['cp:coreProperties'];
        if (core) {
          metadata.creator = this.extractText(core['dc:creator']) || metadata.creator;
          metadata.created = this.extractText(core['dcterms:created']) || metadata.created;
          metadata.modified = this.extractText(core['dcterms:modified']) || metadata.modified;
          metadata.title = this.extractText(core['dc:title']) || undefined;
          metadata.subject = this.extractText(core['dc:subject']) || undefined;
        }
      }
    } catch {}

    return metadata;
  }
}

export const parseExcel = (buffer: Uint8Array): ParsedWorkbook => {
  const reader = new ExcelReader();
  return reader.parseFromBuffer(buffer);
};
