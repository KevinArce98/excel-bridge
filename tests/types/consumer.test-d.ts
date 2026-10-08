import { describe, expectTypeOf, it } from 'vitest';
import * as library from '../../src';
import {
  ExcelBridge,
  ExcelReader,
  ExcelWriter,
  Workbook,
  objectsToSheet,
  objectsToStreamingSheet,
  sheetToObjects,
} from '../../src';
import type {
  CellStyle,
  CellValue,
  ExcelBridgeError,
  ExcelData,
  ExcelErrorValue,
  FormulaCell,
  ObjectCellValue,
  ParsedBorder,
  ParsedCell,
  ParsedRow,
  ParsedSheet,
  SheetLayout,
  SheetOptions,
} from '../../src';

describe('writer input', () => {
  it('accepts literal rows and style records', () => {
    const rows = [['a', 1, true, new Date(), null, { formula: 'A2*2' }]];
    const styles: Record<string, CellStyle> = { '0-0': { bold: true, border: true } };
    new ExcelWriter().createWorkbookBuffer([{ data: rows, styles }]);
  });

  it('keeps the input widening one-way', () => {
    expectTypeOf<string>().toMatchTypeOf<CellValue>();
    expectTypeOf<FormulaCell>().toMatchTypeOf<CellValue>();
    // @ts-expect-error an object without formula, text or error is not a cell
    const bad: CellValue = { result: 1 };
    void bad;
  });

  it('accepts a validation without the removed options field', () => {
    const validation: ExcelData['validations'] = [{ range: 'A1', type: 'custom', formula1: '1' }];
    void validation;
  });
});

describe('ParsedCell is a union on type', () => {
  it('types the value of each variant', () => {
    const check = (cell: ParsedCell) => {
      switch (cell.type) {
        case 'string':
          expectTypeOf(cell.value).toEqualTypeOf<string>();
          break;
        case 'number':
          expectTypeOf(cell.value).toEqualTypeOf<number>();
          break;
        case 'boolean':
          expectTypeOf(cell.value).toEqualTypeOf<boolean>();
          break;
        case 'date':
          expectTypeOf(cell.value).toEqualTypeOf<Date>();
          break;
        case 'error':
          expectTypeOf(cell.value).toMatchTypeOf<string>();
          expectTypeOf<ExcelErrorValue>().toMatchTypeOf<typeof cell.value>();
          break;
        case 'empty':
          expectTypeOf(cell.value).toBeNull();
          break;
        default:
          expectTypeOf(cell).toBeNever();
      }
    };
    void check;
  });

  it('keeps coordinates and the formula on every variant', () => {
    expectTypeOf<ParsedCell['coordinate']>().toBeString();
    expectTypeOf<ParsedCell['rowIndex']>().toBeNumber();
    expectTypeOf<ParsedCell['columnIndex']>().toBeNumber();
    expectTypeOf<ParsedCell['formula']>().toEqualTypeOf<string | undefined>();
  });

  it('rejects a value that does not match its type', () => {
    // @ts-expect-error a number cell holds a number
    const bad: ParsedCell = {
      type: 'number',
      value: 'x',
      coordinate: 'A1',
      rowIndex: 0,
      columnIndex: 0,
    };
    void bad;
  });
});

describe('what the reader returns', () => {
  it('types the rows of a sheet and the sheet list', () => {
    const sheet = new ExcelReader().parseFromBuffer(new Uint8Array()).sheets[0];
    expectTypeOf(sheet).toEqualTypeOf<ParsedSheet>();
    expectTypeOf(sheet.data).toEqualTypeOf<ParsedRow[]>();
    expectTypeOf(ExcelBridge.read).returns.toHaveProperty('sheets');
  });

  it('returns styles that the writer takes back unchanged', () => {
    const sheet = ExcelBridge.read(new Uint8Array()).sheets[0];
    const styles = sheet.styles ?? {};
    const sheetInput: ExcelData = { data: [], styles };
    void sheetInput;
    expectTypeOf(styles['0-0']?.border).toEqualTypeOf<ParsedBorder | undefined>();
    expectTypeOf<ParsedBorder>().toMatchTypeOf<NonNullable<CellStyle['border']>>();
  });

  it('carries the layout of the writer', () => {
    expectTypeOf<ParsedSheet>().toMatchTypeOf<SheetLayout>();
    expectTypeOf<SheetOptions>().toMatchTypeOf<SheetLayout>();
  });
});

describe('Workbook values', () => {
  it('returns cell values, which can be a formula cell', () => {
    const workbook = Workbook.create();
    expectTypeOf(workbook.getCellValue('S', 0, 0)).toEqualTypeOf<CellValue>();
    expectTypeOf(workbook.getSheetData('S')).toEqualTypeOf<CellValue[][]>();
    const value = workbook.getCellValue('S', 0, 0);
    if (typeof value === 'object' && value !== null && !(value instanceof Date)) {
      if (value.formula !== undefined) {
        expectTypeOf(value).toEqualTypeOf<FormulaCell>();
      } else if (value.error !== undefined) {
        expectTypeOf(value.error).toEqualTypeOf<ExcelErrorValue>();
      }
    }
  });
});

describe('removed exports', () => {
  it('no longer exports the low-level helpers', () => {
    expectTypeOf<typeof library>().not.toHaveProperty('StyleManager');
    expectTypeOf<typeof library>().not.toHaveProperty('generateSheetXml');
    expectTypeOf<typeof library>().not.toHaveProperty('generateStylesXml');
    expectTypeOf<typeof library>().not.toHaveProperty('generateColsXml');
    expectTypeOf<typeof library>().not.toHaveProperty('createExcelBlob');
    expectTypeOf<typeof library>().not.toHaveProperty('extractExcelFiles');
    expectTypeOf<typeof library>().not.toHaveProperty('validateExcelStructure');
    expectTypeOf<typeof library>().not.toHaveProperty('CONTENT_TYPES');
    expectTypeOf<typeof library>().not.toHaveProperty('validateRowIndex');
  });

  it('keeps the entry points', () => {
    expectTypeOf<typeof library>().toHaveProperty('ExcelWriter');
    expectTypeOf<typeof library>().toHaveProperty('ExcelReader');
    expectTypeOf<typeof library>().toHaveProperty('Workbook');
    expectTypeOf<typeof library>().toHaveProperty('ExcelBridge');
    expectTypeOf<typeof library>().toHaveProperty('createExcelWorkbookStream');
    expectTypeOf<typeof library>().toHaveProperty('isExcelError');
  });
});

describe('2.0 object helpers and errors', () => {
  interface Person {
    name: string;
    age: number | null;
  }

  it('types the rows of sheetToObjects by the row type', () => {
    const sheet = {} as ParsedSheet;
    const people = sheetToObjects<Person>(sheet, {
      columns: [
        { key: 'name', header: 'Full name' },
        { key: 'age', parse: value => (typeof value === 'number' ? value : null) },
      ],
    });
    expectTypeOf(people).toEqualTypeOf<Person[]>();
    expectTypeOf(sheetToObjects(sheet)).toEqualTypeOf<Record<string, ObjectCellValue>[]>();
  });

  it('rejects a key that is not in the row type and a parse that returns the wrong type', () => {
    const sheet = {} as ParsedSheet;
    // @ts-expect-error nope is not a key of Person
    sheetToObjects<Person>(sheet, { columns: [{ key: 'nope' }] });
    // @ts-expect-error parse for age must return number | null
    sheetToObjects<Person>(sheet, { columns: [{ key: 'age', parse: () => 'text' }] });
  });

  it('types the columns of objectsToSheet by the row type', () => {
    const rows: Person[] = [{ name: 'Ann', age: 3 }];
    expectTypeOf(objectsToSheet(rows, [{ key: 'name' }, { key: 'age', numberFormat: '0' }])).toEqualTypeOf<ExcelData>();
    // @ts-expect-error nope is not a key of Person
    objectsToSheet(rows, [{ key: 'nope' }]);
    // @ts-expect-error a streaming sheet takes no column style
    objectsToStreamingSheet(rows, [{ key: 'age', style: { bold: true } }]);
  });

  it('lets a switch over the code be exhaustive', () => {
    const describeCode = (error: ExcelBridgeError): string => {
      switch (error.code) {
        case 'INVALID_INPUT':
        case 'INVALID_FILE':
        case 'LIMIT_EXCEEDED':
        case 'UNSUPPORTED':
          return error.code;
      }
    };
    expectTypeOf(describeCode).returns.toBeString();
  });
});
