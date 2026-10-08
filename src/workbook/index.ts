import { invalidInput } from '../core/errors';
import { ExcelReader, ParsedCell, ParsedWorkbook } from '../reader';
import type { ExcelReaderOptions } from '../reader';
import { ExcelWriter } from '../writer';
import { parseRange, formatRange } from '../core/cell-ref';
import { EXCEL_LIMITS } from '../core/date-utils';
import { validateSheetName } from '../core/sheet-name';
import { isExcelError } from '../core/cells';
import { prepareLayout } from '../core/sheet-layout';
import { HYPERLINK_STYLE, prepareHyperlink } from '../core/hyperlinks';
import {
  AutoFilter,
  CellValue,
  ErrorCell,
  TextCell,
  CellValidation,
  CellStyle,
  ConditionalFormat,
  Hyperlink,
  SheetState,
} from '../core/types';
import type { ExcelWriterOptions } from '../writer';

interface WorkbookSheet {
  name: string;
  state?: SheetState;
  data: CellValue[][];
  styles: Record<string, CellStyle>;
  validations: CellValidation[];
  mergeCells: string[];
  conditionalFormats: ConditionalFormat[];
  hyperlinks: Hyperlink[];
  autoFilter?: AutoFilter;
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  rowHeights?: Record<number, number>;
  hiddenRows: Set<number>;
  hiddenColumns: Set<number>;
  autoWidth?: boolean;
  literals: Map<string, LiteralCell>;
}

type LiteralCell = TextCell | ErrorCell;

const setMember = (members: Set<number>, index: number, present: boolean): void => {
  if (present) members.add(index);
  else members.delete(index);
};

const cellToValue = (cell: ParsedCell): CellValue => {
  if (cell.formula !== undefined) return `=${cell.formula}`;
  if (cell.type === 'empty') return null;
  return cell.value;
};

const literalOf = (cell: ParsedCell): LiteralCell | undefined => {
  if (cell.formula !== undefined) return undefined;
  if (cell.type === 'error' && isExcelError(cell.value)) return { error: cell.value };
  if (typeof cell.value === 'string' && cell.value.startsWith('=')) return { text: cell.value };
  return undefined;
};

const loadedText = (literal: LiteralCell): string => literal.error ?? literal.text;

const withLiterals = (sheet: WorkbookSheet): CellValue[][] => {
  if (sheet.literals.size === 0) return sheet.data;
  const data = sheet.data.slice();
  sheet.literals.forEach((literal, key) => {
    const [row, col] = key.split('-').map(Number);
    if (data[row]?.[col] !== loadedText(literal)) return;
    data[row] = data[row].slice();
    data[row][col] = literal;
  });
  return data;
};

const placeRows = (
  rows: ParsedCell[][]
): { data: CellValue[][]; literals: Map<string, LiteralCell> } => {
  const data: CellValue[][] = [];
  const literals = new Map<string, LiteralCell>();
  let next = 0;

  for (const row of rows) {
    const declared = row[0]?.rowIndex ?? -1;
    const index = declared >= next ? declared : next;
    data[index] = row.map((cell, col) => {
      const literal = literalOf(cell);
      if (literal !== undefined) literals.set(`${index}-${col}`, literal);
      return cellToValue(cell);
    });
    next = index + 1;
  }

  return { data, literals };
};

const canonicalHyperlink = (link: Hyperlink): Hyperlink => ({
  ...link,
  range: prepareHyperlink(link).ref,
});

const canonicalAutoFilter = (autoFilter: AutoFilter): AutoFilter => ({
  range: formatRange(parseRange(autoFilter.range)),
});

const acceptOrDrop = <T>(value: T, canonicalize: (value: T) => T): T[] => {
  try {
    return [canonicalize(value)];
  } catch {
    return [];
  }
};

const loadHyperlinks = (links: Hyperlink[] = []): Hyperlink[] => {
  const byRange = new Map<string, Hyperlink>();

  for (const link of links.flatMap(link => acceptOrDrop(link, canonicalHyperlink))) {
    if (byRange.size === EXCEL_LIMITS.MAX_HYPERLINKS) break;
    if (!byRange.has(link.range)) byRange.set(link.range, link);
  }

  return [...byRange.values()];
};

export interface WorkbookMetadata {
  creator?: string;
  title?: string;
  subject?: string;
}

export class Workbook {
  private sheets: WorkbookSheet[] = [];
  private metadata: WorkbookMetadata = {};

  private constructor() {}

  static create(): Workbook {
    return new Workbook();
  }

  static fromBuffer(buffer: Uint8Array, options?: ExcelReaderOptions): Workbook {
    const reader = new ExcelReader(options);
    return Workbook.fromParsed(reader.parseFromBuffer(buffer));
  }

  static async fromFile(file: File, options?: ExcelReaderOptions): Promise<Workbook> {
    const reader = new ExcelReader(options);
    return Workbook.fromParsed(await reader.parseFromFile(file));
  }

  private static fromParsed(parsed: ParsedWorkbook): Workbook {
    const workbook = new Workbook();
    workbook.metadata = {
      creator: parsed.metadata.creator,
      title: parsed.metadata.title,
      subject: parsed.metadata.subject,
    };
    workbook.sheets = parsed.sheets.map(sheet => {
      const placed = placeRows(sheet.data);
      return {
        name: sheet.name,
        ...(sheet.state ? { state: sheet.state } : {}),
        data: placed.data,
        literals: placed.literals,
        styles: sheet.styles ?? {},
        validations: sheet.validations.map(validation => ({ ...validation })),
        mergeCells: sheet.mergeCells ?? [],
        conditionalFormats: sheet.conditionalFormats ?? [],
        hyperlinks: loadHyperlinks(sheet.hyperlinks),
        autoFilter: sheet.autoFilter
          ? acceptOrDrop(sheet.autoFilter, canonicalAutoFilter)[0]
          : undefined,
        freezePane: sheet.freezePane,
        columnWidths: sheet.columnWidths,
        rowHeights: sheet.rowHeights,
        hiddenRows: new Set(sheet.hiddenRows),
        hiddenColumns: new Set(sheet.hiddenColumns),
      };
    });
    return workbook;
  }

  getSheetNames(): string[] {
    return this.sheets.map(sheet => sheet.name);
  }

  getMetadata(): WorkbookMetadata {
    return { ...this.metadata };
  }

  setMetadata(metadata: WorkbookMetadata): void {
    this.metadata = { ...this.metadata, ...metadata };
  }

  private findSheet(name: string): WorkbookSheet {
    const sheet = this.sheets.find(s => s.name === name);
    if (!sheet) {
      throw invalidInput(`Sheet "${name}" not found`);
    }
    return sheet;
  }

  addSheet(name: string, data: CellValue[][] = []): void {
    validateSheetName(name);
    if (this.sheets.some(s => s.name.toLowerCase() === name.toLowerCase())) {
      throw invalidInput(`Sheet "${name}" already exists`);
    }
    this.sheets.push({
      name,
      data,
      styles: {},
      validations: [],
      mergeCells: [],
      conditionalFormats: [],
      hyperlinks: [],
      hiddenRows: new Set(),
      hiddenColumns: new Set(),
      literals: new Map(),
    });
  }

  renameSheet(from: string, to: string): void {
    const sheet = this.findSheet(from);
    validateSheetName(to);
    if (this.sheets.some(s => s !== sheet && s.name.toLowerCase() === to.toLowerCase())) {
      throw invalidInput(`Sheet "${to}" already exists`);
    }
    sheet.name = to;
  }

  removeSheet(name: string): void {
    const index = this.sheets.findIndex(s => s.name === name);
    if (index === -1) {
      throw invalidInput(`Sheet "${name}" not found`);
    }
    this.sheets.splice(index, 1);
  }

  getSheetData(name: string): CellValue[][] {
    return this.findSheet(name).data;
  }

  getCellValue(sheetName: string, row: number, col: number): CellValue {
    return this.findSheet(sheetName).data[row]?.[col] ?? null;
  }

  setCellValue(sheetName: string, row: number, col: number, value: CellValue): void {
    const sheet = this.findSheet(sheetName);
    if (!sheet.data[row]) sheet.data[row] = [];
    sheet.data[row][col] = value;
    sheet.literals.delete(`${row}-${col}`);
  }

  getCellStyle(sheetName: string, row: number, col: number): CellStyle | undefined {
    return this.findSheet(sheetName).styles[`${row}-${col}`];
  }

  setCellStyle(sheetName: string, row: number, col: number, style: CellStyle): void {
    this.findSheet(sheetName).styles[`${row}-${col}`] = style;
  }

  getSheetState(sheetName: string): SheetState {
    return this.findSheet(sheetName).state ?? 'visible';
  }

  setSheetState(sheetName: string, state: SheetState): void {
    const sheet = this.findSheet(sheetName);
    if (state === 'visible') {
      delete sheet.state;
    } else {
      sheet.state = state;
    }
  }

  setMergeCells(sheetName: string, ranges: string[]): void {
    this.findSheet(sheetName).mergeCells = ranges;
  }

  setFreezePane(sheetName: string, pane: { row?: number; col?: number }): void {
    this.findSheet(sheetName).freezePane = pane;
  }

  setColumnWidths(sheetName: string, widths: number[]): void {
    this.findSheet(sheetName).columnWidths = widths;
  }

  setRowHeight(sheetName: string, row: number, height: number | null): void {
    const sheet = this.findSheet(sheetName);
    if (height === null) {
      delete sheet.rowHeights?.[row];
    } else {
      prepareLayout({ rowHeights: { [row]: height } });
      (sheet.rowHeights ??= {})[row] = height;
    }
  }

  getRowHeight(sheetName: string, row: number): number | undefined {
    return this.findSheet(sheetName).rowHeights?.[row];
  }

  setRowHidden(sheetName: string, row: number, hidden: boolean = true): void {
    prepareLayout({ hiddenRows: [row] });
    setMember(this.findSheet(sheetName).hiddenRows, row, hidden);
  }

  isRowHidden(sheetName: string, row: number): boolean {
    return this.findSheet(sheetName).hiddenRows.has(row);
  }

  setColumnHidden(sheetName: string, col: number, hidden: boolean = true): void {
    prepareLayout({ hiddenColumns: [col] });
    setMember(this.findSheet(sheetName).hiddenColumns, col, hidden);
  }

  isColumnHidden(sheetName: string, col: number): boolean {
    return this.findSheet(sheetName).hiddenColumns.has(col);
  }

  setAutoWidth(sheetName: string, enabled: boolean): void {
    this.findSheet(sheetName).autoWidth = enabled;
  }

  addValidation(sheetName: string, validation: CellValidation): void {
    this.findSheet(sheetName).validations.push(validation);
  }

  addConditionalFormat(sheetName: string, format: ConditionalFormat): void {
    this.findSheet(sheetName).conditionalFormats.push(format);
  }

  setAutoFilter(sheetName: string, autoFilter: AutoFilter): void {
    this.findSheet(sheetName).autoFilter = canonicalAutoFilter(autoFilter);
  }

  getAutoFilter(sheetName: string): AutoFilter | undefined {
    const autoFilter = this.findSheet(sheetName).autoFilter;
    return autoFilter ? { ...autoFilter } : undefined;
  }

  removeAutoFilter(sheetName: string): void {
    const sheet = this.findSheet(sheetName);
    if (sheet.autoFilter) {
      const { start, end } = parseRange(sheet.autoFilter.range);
      sheet.hiddenRows.forEach(row => {
        if (row > start.row && row <= end.row) sheet.hiddenRows.delete(row);
      });
    }
    delete sheet.autoFilter;
  }

  setHyperlink(sheetName: string, hyperlink: Hyperlink): void {
    const sheet = this.findSheet(sheetName);
    const stored = canonicalHyperlink(hyperlink);
    const index = sheet.hyperlinks.findIndex(link => link.range === stored.range);

    if (index === -1) {
      sheet.hyperlinks.push(stored);
    } else {
      sheet.hyperlinks[index] = stored;
    }
  }

  getHyperlinks(sheetName: string): Hyperlink[] {
    return this.findSheet(sheetName).hyperlinks.map(link => ({ ...link }));
  }

  removeHyperlink(sheetName: string, range: string): void {
    const sheet = this.findSheet(sheetName);
    const parsed = parseRange(range);
    const ref = formatRange(parsed);
    const index = sheet.hyperlinks.findIndex(link => link.range === ref);
    if (index === -1) return;

    sheet.hyperlinks.splice(index, 1);

    const anchor = `${parsed.start.row}-${parsed.start.col}`;
    const style = sheet.styles[anchor];
    if (!style) return;

    const remaining: CellStyle = { ...style };
    if (remaining.underline === true) delete remaining.underline;
    if (remaining.color?.toUpperCase() === HYPERLINK_STYLE.color) delete remaining.color;

    if (Object.keys(remaining).length > 0) {
      sheet.styles[anchor] = remaining;
    } else {
      delete sheet.styles[anchor];
    }
  }

  private toExcelData() {
    return this.sheets.map(sheet => ({
      data: withLiterals(sheet),
      validations: sheet.validations,
      styles: sheet.styles,
      mergeCells: sheet.mergeCells,
      conditionalFormats: sheet.conditionalFormats,
      hyperlinks: sheet.hyperlinks,
      options: {
        name: sheet.name,
        state: sheet.state,
        freezePane: sheet.freezePane,
        columnWidths: sheet.columnWidths,
        rowHeights: sheet.rowHeights,
        hiddenRows: [...sheet.hiddenRows],
        hiddenColumns: [...sheet.hiddenColumns],
        autoWidth: sheet.autoWidth,
        autoFilter: sheet.autoFilter,
      },
    }));
  }

  toBuffer(options?: Partial<ExcelWriterOptions>): Uint8Array {
    const writer = new ExcelWriter({ ...this.metadata, ...options });
    return writer.createWorkbookBuffer(this.toExcelData());
  }

  toBlob(options?: Partial<ExcelWriterOptions>): Blob {
    const writer = new ExcelWriter({ ...this.metadata, ...options });
    return writer.createWorkbook(this.toExcelData());
  }
}
