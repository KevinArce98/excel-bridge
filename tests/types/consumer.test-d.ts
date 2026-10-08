import { describe, expectTypeOf, it } from 'vitest';
import {
  ExcelWriter,
  StyleManager,
  objectsToSheet,
  objectsToStreamingSheet,
  sheetToObjects,
} from '../../src';
import type {
  CellStyle,
  CellValue,
  ExcelBridgeError,
  ExcelData,
  ObjectCellValue,
  ParsedSheet,
} from '../../src';

describe('1.5 consumer code keeps compiling', () => {
  it('accepts literal rows and style records', () => {
    const rows = [['a', 1, true, new Date(), null]];
    const styles: Record<string, CellStyle> = { '0-0': { bold: true, border: true } };
    new ExcelWriter().createWorkbookBuffer([{ data: rows, styles }]);
    expectTypeOf(new StyleManager().getStyleId({ bold: true })).toBeNumber();
  });

  it('keeps the input widening one-way', () => {
    expectTypeOf<string>().toMatchTypeOf<CellValue>();
    // @ts-expect-error an object without formula, text or error is not a cell
    const bad: CellValue = { result: 1 };
    void bad;
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
