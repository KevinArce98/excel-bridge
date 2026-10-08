import { XMLParser } from 'fast-xml-parser';
import { extractParts, validateExcelStructure } from '../core/zip-manager';
import type { PartBudget } from '../core/zip-manager';
import { ExcelBridgeError, invalidInput, limitExceeded } from '../core/errors';
import type { ReaderLimitName } from '../core/errors';
import {
  EXCEL_LIMITS,
  excelSerialToDate,
  isDate,
  isDateNumFmtId,
  validateColIndex,
  validateRowIndex,
} from '../core/date-utils';
import { isRowHeight } from '../core/row-height';
import { BORDER_SIDES, BORDER_STYLES } from '../core/borders';
import type { BorderLine, BorderLines, BorderStyleName, ParsedBorder } from '../core/borders';
import type {
  AutoFilter,
  CellStyle,
  CellValidation,
  ExcelErrorValue,
  ConditionalFormat,
  ConditionalFormatStyle,
  ConditionalFormatOperator,
  DataValidationOperator,
  DataValidationType,
  Hyperlink,
  SheetLayout,
  SheetState,
} from '../core/types';

interface ParsedCellBase {
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
  formula?: string;
}

export interface ParsedStringCell extends ParsedCellBase {
  type: 'string';
  value: string;
}

export interface ParsedNumberCell extends ParsedCellBase {
  type: 'number';
  value: number;
}

export interface ParsedBooleanCell extends ParsedCellBase {
  type: 'boolean';
  value: boolean;
}

export interface ParsedDateCell extends ParsedCellBase {
  type: 'date';
  value: Date;
}

export interface ParsedErrorCell extends ParsedCellBase {
  type: 'error';
  value: ExcelErrorValue | (string & {});
}

export interface ParsedEmptyCell extends ParsedCellBase {
  type: 'empty';
  value: null;
}

export type ParsedCell =
  | ParsedStringCell
  | ParsedNumberCell
  | ParsedBooleanCell
  | ParsedDateCell
  | ParsedErrorCell
  | ParsedEmptyCell;

export type ParsedRow = ParsedCell[];

export interface ParsedSheet extends SheetLayout {
  name: string;
  data: ParsedRow[];
  validations: CellValidation[];
  state?: Exclude<SheetState, 'visible'>;
  styles?: Record<string, CellStyle<ParsedBorder>>;
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  autoFilter?: AutoFilter;
  hyperlinks?: Hyperlink[];
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

const isBorderStyleName = (style: unknown): style is BorderStyleName =>
  BORDER_STYLES.includes(style as BorderStyleName);

type RawBorder = Partial<Record<(typeof BORDER_SIDES)[number], RawBorderSide>> | undefined;

interface RawBorderSide {
  style?: string;
  color?: { rgb?: string };
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
  borders: (ParsedBorder | undefined)[];
  customFormats: Record<number, string>;
  cellXfs: DecodedXf[];
  dateStyles: Set<number>;
  dxfs: ConditionalFormatStyle[];
}

const VALIDATION_TYPES = new Set([
  'list',
  'whole',
  'decimal',
  'textLength',
  'date',
  'time',
  'custom',
]);

const VALIDATION_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
]);

const toArray = <T>(value: T | T[] | undefined): T[] => {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
};

export interface ExcelReaderOptions {
  maxCells?: number;
  maxPartBytes?: number;
  maxTotalBytes?: number;
  maxSheets?: number;
}

export const DEFAULT_READER_LIMITS: Required<ExcelReaderOptions> = {
  maxCells: 5_000_000,
  maxPartBytes: 268_435_456,
  maxTotalBytes: 536_870_912,
  maxSheets: Infinity,
};

const resolveLimits = (options: ExcelReaderOptions = {}): Required<ExcelReaderOptions> => {
  const limits = { ...DEFAULT_READER_LIMITS };
  (Object.keys(limits) as ReaderLimitName[]).forEach(name => {
    const value = options[name];
    if (value === undefined) return;
    if (!(value > 0)) throw invalidInput(`${name} must be a number above 0, or Infinity`);
    limits[name] = value;
  });
  return limits;
};

const STRING_TYPES = new Set(['s', 'str', 'e', 'b']);

const withFormula = <T extends ParsedCell>(cell: T, formula: string): T => {
  if (formula) cell.formula = formula;
  return cell;
};

const isFlag = (value: unknown): boolean => value === '1' || value === 'true';

const BUILT_IN_DATE_FORMATS: Record<number, string> = {
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
};

const STRUCTURAL_PARTS = new Set([
  '[Content_Types].xml',
  '_rels/.rels',
  'xl/workbook.xml',
  'xl/_rels/workbook.xml.rels',
  'xl/sharedStrings.xml',
  'xl/styles.xml',
  'docProps/app.xml',
  'docProps/core.xml',
]);

const relsPathFor = (partPath: string): string =>
  partPath.replace(/[^/]+$/, name => `_rels/${name}.rels`);

const loadSheetParts = (
  buffer: Uint8Array,
  files: Record<string, string>,
  sheetPaths: Array<string | undefined>,
  budget: PartBudget
): void => {
  const wanted = new Set<string>();
  for (const path of sheetPaths) {
    if (path) {
      wanted.add(path);
      wanted.add(relsPathFor(path));
    }
  }
  Object.assign(
    files,
    extractParts(buffer, path => wanted.has(path), budget)
  );
};

export class ExcelReader {
  private parser: XMLParser;
  private limits: Required<ExcelReaderOptions>;

  constructor(options?: ExcelReaderOptions) {
    this.limits = resolveLimits(options);
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
      textNodeName: '#text',
      parseAttributeValue: false,
      parseTagValue: false,
      trimValues: false,
    });
  }

  async parseFromFile(file: File): Promise<ParsedWorkbook> {
    const buffer = new Uint8Array(await file.arrayBuffer());
    return this.parseFromBuffer(buffer);
  }

  parseFromBuffer(buffer: Uint8Array): ParsedWorkbook {
    try {
      const budget: PartBudget = { ...this.limits, total: 0 };
      const files = extractParts(buffer, path => STRUCTURAL_PARTS.has(path), budget);

      if (!files['xl/workbook.xml']) {
        throw new ExcelBridgeError('INVALID_FILE', 'Invalid Excel file structure');
      }

      const workbook = this.parser.parse(files['xl/workbook.xml']);
      const sharedStrings = this.parseSharedStrings(files);
      const styleSheet = this.parseStyleSheet(files);
      const relMap = this.parseWorkbookRels(files);

      const sheets: ParsedSheet[] = [];
      const sheetElements = toArray(workbook.workbook?.sheets?.sheet);
      if (sheetElements.length > this.limits.maxSheets) {
        throw limitExceeded(
          'maxSheets',
          this.limits.maxSheets,
          `Workbook has ${sheetElements.length} sheets`
        );
      }
      const sheetPaths = sheetElements.map((sheetElement, index) =>
        this.resolveSheetPath(sheetElement, relMap, index)
      );
      loadSheetParts(buffer, files, sheetPaths, budget);

      if (!validateExcelStructure(files)) {
        throw new ExcelBridgeError('INVALID_FILE', 'Invalid Excel file structure');
      }

      const cellBudget = { cells: 0 };

      sheetElements.forEach((sheetElement, index) => {
        const sheetName = sheetElement.name ?? `Sheet${index + 1}`;
        const sheetPath = sheetPaths[index];

        if (sheetPath && files[sheetPath]) {
          const sheetData = this.parseSheet(
            files[sheetPath],
            sharedStrings,
            styleSheet,
            cellBudget,
            files[relsPathFor(sheetPath)]
          );
          const state = sheetElement.state;
          sheets.push({
            name: String(sheetName),
            ...(state === 'hidden' || state === 'veryHidden' ? { state } : {}),
            ...sheetData,
          });
        }
      });

      return {
        sheets,
        metadata: this.extractMetadata(files),
      };
    } catch (error) {
      throw new ExcelBridgeError(
        error instanceof ExcelBridgeError && error.code === 'LIMIT_EXCEEDED'
          ? 'LIMIT_EXCEEDED'
          : 'INVALID_FILE',
        `Failed to parse Excel file: ${error instanceof Error ? error.message : 'Unknown error'}`,
        {
          cause: error,
          limit: error instanceof ExcelBridgeError ? error.limit : undefined,
        }
      );
    }
  }

  private parseRelationships(relsXml?: string): Record<string, string> {
    const map: Record<string, string> = {};
    if (!relsXml) return map;

    try {
      const parsed = this.parser.parse(relsXml);
      const rels = toArray(parsed.Relationships?.Relationship);
      for (const rel of rels) {
        if (!rel.Id || !rel.Target) continue;
        map[String(rel.Id)] = String(rel.Target);
      }
    } catch {}

    return map;
  }

  private parseWorkbookRels(files: Record<string, string>): Record<string, string> {
    const map = this.parseRelationships(files['xl/_rels/workbook.xml.rels']);

    for (const [id, target] of Object.entries(map)) {
      map[id] = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    }

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
    if (typeof item === 'string') return item;

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
      dxfs: [],
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

      result.borders = toArray(styleSheet.borders?.border).map(border => this.decodeBorder(border));

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
                wrapText: isFlag(alignment.wrapText),
              }
            : undefined,
        };
      });

      result.cellXfs.forEach((xf, index) => {
        if (isDateNumFmtId(xf.numFmtId, result.customFormats)) {
          result.dateStyles.add(index);
        }
      });

      result.dxfs = toArray(styleSheet.dxfs?.dxf).map((dxf: any) => {
        const style: ConditionalFormatStyle = {};
        const font = dxf?.font;
        if (font) {
          if (font.b !== undefined) style.bold = true;
          if (font.i !== undefined) style.italic = true;
          if (font.color?.rgb !== undefined) style.color = this.argbToHex(String(font.color.rgb));
        }
        const bgColor = dxf?.fill?.patternFill?.bgColor?.rgb;
        if (bgColor !== undefined) style.background = this.argbToHex(String(bgColor));
        return style;
      });
    } catch {}

    return result;
  }

  private decodeBorder(border: RawBorder): ParsedBorder | undefined {
    const lines: BorderLines = {};
    let plain = 0;
    for (const side of BORDER_SIDES) {
      const { style, color } = border?.[side] ?? {};
      if (!isBorderStyleName(style)) continue;
      const line: BorderLine = { style };
      const hex = color?.rgb === undefined ? '#000000' : this.argbToHex(String(color.rgb));
      if (hex.toUpperCase() !== '#000000') line.color = hex;
      lines[side] = line;
      if (style === 'thin' && !line.color) plain++;
    }
    return plain === 4 ? true : Object.keys(lines).length > 0 ? lines : undefined;
  }

  private argbToHex(argb: string): string {
    const hex = argb.length === 8 ? argb.slice(2) : argb;
    return `#${hex}`;
  }

  private decodeCellStyle(
    styleIndex: number | undefined,
    styleSheet: StyleSheetData
  ): CellStyle<ParsedBorder> | undefined {
    if (styleIndex === undefined) return undefined;

    const xf = styleSheet.cellXfs[styleIndex];
    if (!xf) return undefined;

    const numberFormat =
      styleSheet.customFormats[xf.numFmtId] ?? BUILT_IN_DATE_FORMATS[xf.numFmtId];
    const isPlainDate =
      styleSheet.dateStyles.has(styleIndex) &&
      !numberFormat &&
      !xf.fontId &&
      !xf.fillId &&
      !xf.borderId &&
      !xf.alignment;
    if (isPlainDate) return undefined;

    const font = styleSheet.fonts[xf.fontId];
    const fill = styleSheet.fills[xf.fillId];
    const border = structuredClone(styleSheet.borders[xf.borderId]);

    const style: CellStyle<ParsedBorder> = {};

    if (font?.bold) style.bold = true;
    if (font?.italic) style.italic = true;
    if (font?.underline) style.underline = true;
    if (font?.color) style.color = this.argbToHex(font.color);
    if (font?.size) style.fontSize = font.size;
    if (font?.name) style.fontName = font.name;

    if (fill?.patternType === 'solid' && fill.fgColor) {
      style.background = this.argbToHex(fill.fgColor);
    }

    if (border) style.border = border;

    if (xf.alignment) {
      if (xf.alignment.horizontal) style.align = xf.alignment.horizontal;
      if (xf.alignment.vertical) style.verticalAlign = xf.alignment.vertical;
      if (xf.alignment.wrapText) style.wrapText = true;
    }

    if (numberFormat) style.numberFormat = numberFormat;

    return Object.keys(style).length > 0 ? style : undefined;
  }

  private parseSheet(
    sheetXml: string,
    sharedStrings: string[],
    styleSheet: StyleSheetData,
    cellBudget: { cells: number },
    relsXml?: string
  ): Omit<ParsedSheet, 'name'> {
    const parsed = this.parser.parse(sheetXml);
    const worksheet = parsed.worksheet;

    const rows = toArray(worksheet?.sheetData?.row);
    const validations = toArray(worksheet?.dataValidations?.dataValidation);

    const parsedValidations = validations
      .map((validation: any) => this.parseValidation(validation))
      .filter((validation): validation is CellValidation => validation !== undefined);

    const data: ParsedRow[] = [];
    const styles: Record<string, CellStyle<ParsedBorder>> = {};
    const rowHeights: Record<number, number> = {};
    const hiddenRows: number[] = [];
    let previousRow = -1;

    for (const rowElement of rows) {
      const declaredRow = parseInt(rowElement.r, 10) - 1;
      const rowIndex = declaredRow >= 0 ? declaredRow : previousRow + 1;
      if (rowIndex >= EXCEL_LIMITS.MAX_ROWS) validateRowIndex(rowIndex);
      previousRow = rowIndex;

      const height = Number(rowElement.ht);
      if (isFlag(rowElement.customHeight) && isRowHeight(height)) rowHeights[rowIndex] = height;
      if (isFlag(rowElement.hidden)) hiddenRows.push(rowIndex);

      const cells = toArray(rowElement.c);
      if (cells.length === 0) continue;

      const rowData: ParsedRow = data[rowIndex] ?? [];
      const widthBefore = rowData.length;
      let previousColumn = -1;

      for (const cell of cells) {
        const parsedCell = this.parseCell(
          cell,
          rowIndex,
          previousColumn,
          sharedStrings,
          styleSheet
        );
        previousColumn = parsedCell.columnIndex;
        rowData[previousColumn] = parsedCell;

        const styleIndex = cell.s !== undefined ? Number(cell.s) : undefined;
        const decodedStyle = this.decodeCellStyle(styleIndex, styleSheet);
        if (decodedStyle) {
          styles[`${rowIndex}-${previousColumn}`] = decodedStyle;
        }
      }

      cellBudget.cells += rowData.length - widthBefore;
      if (cellBudget.cells > this.limits.maxCells) {
        throw limitExceeded(
          'maxCells',
          this.limits.maxCells,
          `Workbook has at least ${cellBudget.cells} cells, counting the empty cells that pad rows`
        );
      }

      for (let c = 0; c < rowData.length; c++) {
        rowData[c] ??= {
          value: null,
          type: 'empty',
          coordinate: `${this.columnIndexToLetter(c)}${rowIndex + 1}`,
          rowIndex,
          columnIndex: c,
        };
      }

      data[rowIndex] = rowData;
    }

    const mergeCells = toArray(worksheet?.mergeCells?.mergeCell)
      .map((m: any) => (m?.ref !== undefined ? String(m.ref) : undefined))
      .filter((ref): ref is string => ref !== undefined);

    const freezePane = this.parseFreezePane(worksheet);
    const columns = this.parseColumns(worksheet);
    const conditionalFormats = this.parseConditionalFormats(worksheet, styleSheet);
    const autoFilterRef = worksheet?.autoFilter?.ref;
    const hyperlinks = this.parseHyperlinks(worksheet, relsXml);

    return {
      data,
      validations: parsedValidations,
      ...(Object.keys(styles).length > 0 ? { styles } : {}),
      ...(mergeCells.length > 0 ? { mergeCells } : {}),
      ...(freezePane ? { freezePane } : {}),
      ...columns,
      ...(Object.keys(rowHeights).length > 0 ? { rowHeights } : {}),
      ...(hiddenRows.length > 0 ? { hiddenRows } : {}),
      ...(conditionalFormats.length > 0 ? { conditionalFormats } : {}),
      ...(autoFilterRef !== undefined ? { autoFilter: { range: String(autoFilterRef) } } : {}),
      ...(hyperlinks.length > 0 ? { hyperlinks } : {}),
    };
  }

  private parseValidation(validation: any): CellValidation | undefined {
    const type = validation?.type;
    if (!VALIDATION_TYPES.has(type) || validation.sqref === undefined) return undefined;

    const formula1 =
      validation.formula1 !== undefined ? this.extractText(validation.formula1) : undefined;
    const formula2 =
      validation.formula2 !== undefined ? this.extractText(validation.formula2) : undefined;
    const base = {
      range: String(validation.sqref),
      type: type as DataValidationType,
      allowBlank: isFlag(validation.allowBlank),
    };

    if (type === 'list') return formula1 === undefined ? undefined : { ...base, formula1 };

    return {
      ...base,
      ...(VALIDATION_OPERATORS.has(validation.operator)
        ? { operator: validation.operator as DataValidationOperator }
        : {}),
      ...(formula1 !== undefined ? { formula1 } : {}),
      ...(formula2 !== undefined ? { formula2 } : {}),
    };
  }

  private parseHyperlinks(worksheet: any, relsXml?: string): Hyperlink[] {
    const entries = toArray(worksheet?.hyperlinks?.hyperlink);
    if (entries.length === 0) return [];

    const relationshipId = (entry: any) => entry?.['r:id'] ?? entry?.id;
    const targets = entries.some(entry => relationshipId(entry) !== undefined)
      ? this.parseRelationships(relsXml)
      : {};

    const links: Hyperlink[] = [];
    for (const entry of entries) {
      if (entry?.ref === undefined) continue;

      const range = String(entry.ref);
      const location = entry.location !== undefined ? String(entry.location) : undefined;
      const extras = {
        ...(entry.tooltip !== undefined ? { tooltip: String(entry.tooltip) } : {}),
        ...(entry.display !== undefined ? { display: String(entry.display) } : {}),
      };
      const rId = relationshipId(entry);

      if (rId !== undefined) {
        const target = targets[String(rId)];
        if (target === undefined) continue;
        links.push({ range, url: location ? `${target}#${location}` : target, ...extras });
      } else if (location) {
        links.push({ range, location, ...extras });
      }
    }

    return links;
  }

  private parseConditionalFormats(worksheet: any, styleSheet: StyleSheetData): ConditionalFormat[] {
    const formats: ConditionalFormat[] = [];

    for (const block of toArray(worksheet?.conditionalFormatting)) {
      const range = block?.sqref !== undefined ? String(block.sqref) : '';
      if (!range) continue;

      for (const rule of toArray(block?.cfRule)) {
        const cf = this.parseCfRule(rule, range, styleSheet);
        if (cf) formats.push(cf);
      }
    }

    return formats;
  }

  private parseCfRule(
    rule: any,
    range: string,
    styleSheet: StyleSheetData
  ): ConditionalFormat | undefined {
    const type = rule?.type;

    if (type === 'colorScale') {
      const colors = toArray(rule?.colorScale?.color)
        .map((c: any) => (c?.rgb !== undefined ? this.argbToHex(String(c.rgb)) : undefined))
        .filter((c): c is string => c !== undefined);
      if (colors.length === 2 || colors.length === 3) {
        return {
          type: 'colorScale',
          range,
          colors: colors as [string, string] | [string, string, string],
        };
      }
      return undefined;
    }

    const style = styleSheet.dxfs[Number(rule?.dxfId ?? -1)] ?? {};

    if (type === 'expression') {
      return { type: 'expression', range, formula: this.extractText(rule?.formula), style };
    }

    if (type === 'cellIs') {
      const formulas = toArray(rule?.formula).map((f: any) => this.parseCfValue(f));
      return {
        type: 'cellValue',
        range,
        operator: rule?.operator as ConditionalFormatOperator,
        value: formulas[0],
        ...(formulas.length > 1 ? { value2: formulas[1] } : {}),
        style,
      };
    }

    return undefined;
  }

  private parseCfValue(raw: any): number | string {
    const value = raw && typeof raw === 'object' && raw['#text'] !== undefined ? raw['#text'] : raw;
    const text = String(value ?? '');
    if (text.startsWith('"') && text.endsWith('"')) return text.slice(1, -1);
    const num = Number(text);
    return text !== '' && !Number.isNaN(num) ? num : text;
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

  private parseColumns(worksheet: any): Pick<ParsedSheet, 'columnWidths' | 'hiddenColumns'> {
    const cols = toArray(worksheet?.cols?.col);
    if (cols.length === 0) return {};

    const widths: number[] = [];
    const hidden = new Set<number>();
    cols.forEach((col: any) => {
      const min = Math.max(Number(col.min) - 1, 0);
      const max = Math.min(Number(col.max) - 1, EXCEL_LIMITS.MAX_COLS - 1);
      const width = Number(col.width);
      for (let c = min; c <= max; c++) {
        if (Number.isFinite(width)) widths[c] = width;
        if (isFlag(col.hidden)) hidden.add(c);
      }
    });

    return {
      columnWidths: widths,
      ...(hidden.size > 0 ? { hiddenColumns: [...hidden].sort((a, b) => a - b) } : {}),
    };
  }

  private parseCell(
    cell: any,
    rowIndex: number,
    previousColumn: number,
    sharedStrings: string[],
    styleSheet: StyleSheetData
  ): ParsedCell {
    const declaredColumn = this.columnLetterToIndex(String(cell.r ?? '').replace(/\d+/g, ''));
    const columnIndex = declaredColumn >= 0 ? declaredColumn : previousColumn + 1;
    if (columnIndex >= EXCEL_LIMITS.MAX_COLS) validateColIndex(columnIndex);
    const coordinate = `${this.columnIndexToLetter(columnIndex)}${rowIndex + 1}`;
    const styleIndex = cell.s !== undefined ? Number(cell.s) : undefined;
    const formula = cell.f === undefined ? '' : this.extractText(cell.f);

    if (cell.t === 'inlineStr' || (cell.t === undefined && cell.is !== undefined)) {
      return withFormula(
        {
          coordinate,
          rowIndex,
          columnIndex,
          type: 'string',
          value: this.extractStringItem(cell.is),
        },
        formula
      );
    }

    const raw = cell.v;
    if (raw === undefined || (raw === '' && !STRING_TYPES.has(cell.t))) {
      return withFormula(
        { coordinate, rowIndex, columnIndex, type: 'empty', value: null },
        formula
      );
    }

    if (cell.t === 'b')
      return withFormula(
        { coordinate, rowIndex, columnIndex, type: 'boolean', value: raw === '1' },
        formula
      );
    if (cell.t === 's') {
      return withFormula(
        {
          coordinate,
          rowIndex,
          columnIndex,
          type: 'string',
          value: sharedStrings[parseInt(raw, 10)] ?? '',
        },
        formula
      );
    }
    if (cell.t === 'str')
      return withFormula(
        { coordinate, rowIndex, columnIndex, type: 'string', value: String(raw) },
        formula
      );
    if (cell.t === 'e')
      return withFormula(
        { coordinate, rowIndex, columnIndex, type: 'error', value: String(raw) },
        formula
      );

    const num = parseFloat(raw);
    if (!Number.isFinite(num))
      return withFormula(
        { coordinate, rowIndex, columnIndex, type: 'error', value: '#NUM!' },
        formula
      );

    if (styleIndex !== undefined && styleSheet.dateStyles.has(styleIndex)) {
      const date = excelSerialToDate(num);
      if (isDate(date))
        return withFormula(
          { coordinate, rowIndex, columnIndex, type: 'date', value: date },
          formula
        );
    }
    return withFormula({ coordinate, rowIndex, columnIndex, type: 'number', value: num }, formula);
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

export const parseExcel = (buffer: Uint8Array, options?: ExcelReaderOptions): ParsedWorkbook => {
  const reader = new ExcelReader(options);
  return reader.parseFromBuffer(buffer);
};
