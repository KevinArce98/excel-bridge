import { ExcelReader, ParsedCell, ParsedWorkbook } from '../reader';
import { ExcelWriter } from '../writer';
import { CellValue, CellValidation, CellStyle, ConditionalFormat } from '../core/types';
import type { ExcelWriterOptions } from '../writer';

interface WorkbookSheet {
  name: string;
  data: CellValue[][];
  styles: Record<string, CellStyle>;
  validations: CellValidation[];
  mergeCells: string[];
  conditionalFormats: ConditionalFormat[];
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  autoWidth?: boolean;
}

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

  static fromBuffer(buffer: Uint8Array): Workbook {
    const reader = new ExcelReader();
    return Workbook.fromParsed(reader.parseFromBuffer(buffer));
  }

  static async fromFile(file: File): Promise<Workbook> {
    const reader = new ExcelReader();
    return Workbook.fromParsed(await reader.parseFromFile(file));
  }

  private static fromParsed(parsed: ParsedWorkbook): Workbook {
    const workbook = new Workbook();
    workbook.metadata = {
      creator: parsed.metadata.creator,
      title: parsed.metadata.title,
      subject: parsed.metadata.subject,
    };
    workbook.sheets = parsed.sheets.map(sheet => ({
      name: sheet.name,
      data: sheet.data.map(row => row.map(Workbook.cellToValue)),
      styles: sheet.styles ?? {},
      validations: sheet.validations.map(v => ({ range: v.range, options: v.options })),
      mergeCells: sheet.mergeCells ?? [],
      conditionalFormats: sheet.conditionalFormats ?? [],
      freezePane: sheet.freezePane,
      columnWidths: sheet.columnWidths,
    }));
    return workbook;
  }

  private static cellToValue(cell: ParsedCell): CellValue {
    if (cell.formula !== undefined) return `=${cell.formula}`;
    if (cell.type === 'empty') return null;
    return cell.value;
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
      throw new Error(`Sheet "${name}" not found`);
    }
    return sheet;
  }

  addSheet(name: string, data: CellValue[][] = []): void {
    if (this.sheets.some(s => s.name === name)) {
      throw new Error(`Sheet "${name}" already exists`);
    }
    this.sheets.push({
      name,
      data,
      styles: {},
      validations: [],
      mergeCells: [],
      conditionalFormats: [],
    });
  }

  removeSheet(name: string): void {
    const index = this.sheets.findIndex(s => s.name === name);
    if (index === -1) {
      throw new Error(`Sheet "${name}" not found`);
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
  }

  getCellStyle(sheetName: string, row: number, col: number): CellStyle | undefined {
    return this.findSheet(sheetName).styles[`${row}-${col}`];
  }

  setCellStyle(sheetName: string, row: number, col: number, style: CellStyle): void {
    this.findSheet(sheetName).styles[`${row}-${col}`] = style;
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

  setAutoWidth(sheetName: string, enabled: boolean): void {
    this.findSheet(sheetName).autoWidth = enabled;
  }

  addValidation(sheetName: string, validation: CellValidation): void {
    this.findSheet(sheetName).validations.push(validation);
  }

  addConditionalFormat(sheetName: string, format: ConditionalFormat): void {
    this.findSheet(sheetName).conditionalFormats.push(format);
  }

  private toExcelData() {
    return this.sheets.map(sheet => ({
      data: sheet.data,
      validations: sheet.validations,
      styles: sheet.styles,
      mergeCells: sheet.mergeCells,
      conditionalFormats: sheet.conditionalFormats,
      options: {
        name: sheet.name,
        freezePane: sheet.freezePane,
        columnWidths: sheet.columnWidths,
        autoWidth: sheet.autoWidth,
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
