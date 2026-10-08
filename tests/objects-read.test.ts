import { describe, it, expect } from 'vitest';
import { ExcelReader, ExcelWriter, objectsToSheet, sheetToObjects } from '../src';
import type { CellValue, ExcelData, ParsedSheet } from '../src';
import { buildXlsx } from './helpers/xlsx';

const sheetOf = (data: CellValue[][], extra: Partial<ExcelData> = {}): ParsedSheet =>
  new ExcelReader().parseFromBuffer(new ExcelWriter().createWorkbookBuffer([{ data, ...extra }]))
    .sheets[0];

const failureOf = (action: () => unknown) => {
  try {
    action();
  } catch (error) {
    return error as Error & { code?: string };
  }
  throw new Error('did not throw');
};

describe('sheetToObjects with the header row as keys', () => {
  it('maps the first row to keys and every later row to an object', () => {
    const sheet = sheetOf([
      ['Name', 'Age'],
      ['Ann', 30],
      ['Bob', null],
    ]);
    expect(sheetToObjects(sheet)).toEqual([
      { Name: 'Ann', Age: 30 },
      { Name: 'Bob', Age: null },
    ]);
  });

  it('returns every kind of cell value', () => {
    const born = new Date(2024, 0, 15);
    const sheet = sheetOf([
      [
        'text',
        'number',
        'flag',
        'date',
        'error',
        'formula',
        'formula without result',
        'equals text',
        'empty',
      ],
      [
        'x',
        1.5,
        true,
        born,
        { error: '#N/A' },
        { formula: 'A2', result: 'x' },
        { formula: 'A2' },
        '=A2',
        null,
      ],
    ]);
    expect(sheetToObjects(sheet)).toEqual([
      {
        text: 'x',
        number: 1.5,
        flag: true,
        date: born,
        error: { error: '#N/A' },
        formula: 'x',
        'formula without result': null,
        'equals text': '=A2',
        empty: null,
      },
    ]);
  });

  it('skips rows without a value in any column, including rows that only carry a style', () => {
    const sheet = sheetOf([['a', 'b'], ['1', '2'], [], [null, ''], ['', null], ['3', '4']], {
      styles: { '3-0': { bold: true }, '4-1': { bold: true } },
    });
    expect(sheetToObjects(sheet)).toEqual([
      { a: '1', b: '2' },
      { a: '3', b: '4' },
    ]);
  });

  it('fills the keys of a row that is shorter than the header with null', () => {
    const sheet = sheetOf([['a', 'b', 'c'], ['1']]);
    expect(sheetToObjects(sheet)).toEqual([{ a: '1', b: null, c: null }]);
  });

  it('returns an empty array for a sheet with no rows or only a header', () => {
    expect(sheetToObjects(sheetOf([['a', 'b']]))).toEqual([]);
    expect(sheetToObjects({ data: [] })).toEqual([]);
  });

  it('trims header text, ignores blank headers and keeps repeated headers apart', () => {
    const sheet = sheetOf([
      [' a ', '', 'a', 'a', 'a_1', null, 7],
      [1, 2, 3, 4, 5, 6, 7],
    ]);
    expect(sheetToObjects(sheet)).toEqual([{ a: 1, a_1: 3, a_2: 4, a_1_1: 5, '7': 7 }]);
  });

  it('ignores a header called __proto__', () => {
    const [row] = sheetToObjects(
      sheetOf([
        ['__proto__', 'a'],
        [{ error: '#N/A' }, 1],
      ])
    );
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(row).toEqual({ a: 1 });
  });
});

describe('sheetToObjects error cells', () => {
  it('returns the seven classic errors as { error } and any other error text as a string', () => {
    const body =
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c><c r="C1" t="inlineStr"><is><t>c</t></is></c></row>' +
      '<row r="2"><c r="A2" t="e"><v>#DIV/0!</v></c><c r="B2" t="e"><v>#SPILL!</v></c><c r="C2"><v>1e999</v></c></row></sheetData>';
    const sheet = new ExcelReader().parseFromBuffer(buildXlsx(body)).sheets[0];
    expect(sheetToObjects(sheet)).toEqual([
      { a: { error: '#DIV/0!' }, b: '#SPILL!', c: { error: '#NUM!' } },
    ]);
  });
});

describe('sheetToObjects headerRow', () => {
  const titled = buildXlsx(
    '<sheetData>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>Report</t></is></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>id</t></is></c><c r="B5" t="inlineStr"><is><t>name</t></is></c></row>' +
      '<row r="7"><c r="A7"><v>1</v></c><c r="B7" t="inlineStr"><is><t>Ann</t></is></c></row>' +
      '<row r="9"><c r="A9"><v>2</v></c><c r="B9" t="inlineStr"><is><t>Bob</t></is></c></row>' +
      '</sheetData>'
  );
  const sheet = new ExcelReader().parseFromBuffer(titled).sheets[0];

  it('finds the header by its row index, not by its position in data', () => {
    expect(sheetToObjects(sheet, { headerRow: 4 })).toEqual([
      { id: 1, name: 'Ann' },
      { id: 2, name: 'Bob' },
    ]);
  });

  it('finds nothing at a row that is not in the file', () => {
    expect(sheetToObjects(sheet, { headerRow: 3 })).toEqual([]);
  });

  it('reads a data array without holes the same way', () => {
    const dense = { data: Object.values(sheet.data) };
    expect(dense.data).toHaveLength(4);
    expect(sheetToObjects(dense, { headerRow: 4 })).toEqual(
      sheetToObjects(sheet, { headerRow: 4 })
    );
  });

  it('reads a sheet without a header row when columns give the positions', () => {
    const rows = sheetToObjects<{ id: number; name: string }>(
      sheetOf([
        [1, 'Ann'],
        [2, 'Bob'],
      ]),
      {
        headerRow: null,
        columns: [
          { key: 'id', column: 0 },
          { key: 'name', column: 1 },
        ],
      }
    );
    expect(rows).toEqual([
      { id: 1, name: 'Ann' },
      { id: 2, name: 'Bob' },
    ]);
  });

  it('without a header row and without positions there is nothing to match', () => {
    const failure = failureOf(() =>
      sheetToObjects(sheetOf([[1]]), { headerRow: null, columns: [{ key: 'id' }] })
    );
    expect(failure.code).toBe('INVALID_INPUT');
    expect(failure.message).toBe('Column "id" needs a column index because there is no header row');
  });

  it('without a header row and without columns there are no keys', () => {
    expect(sheetToObjects(sheetOf([[1]]), { headerRow: null })).toEqual([]);
  });
});

describe('sheetToObjects columns', () => {
  interface Person {
    name: string;
    born: Date | null;
    email: string | null;
  }
  const born = new Date(2000, 4, 6);
  const sheet = sheetOf([
    ['Full name', 'Birth date', 'Ignored'],
    ['Ann', born, 'x'],
    ['Bob', null, 'y'],
  ]);

  it('maps header text to keys, keeps only the listed columns and applies parse', () => {
    const calls: unknown[][] = [];
    const rows = sheetToObjects<Pick<Person, 'name' | 'born'>>(sheet, {
      columns: [
        {
          key: 'born',
          header: 'Birth date',
          parse: (value, row, column) => (
            calls.push([row, column]),
            value instanceof Date ? value : null
          ),
        },
        { key: 'name', header: 'Full name' },
      ],
    });

    expect(rows).toEqual([
      { born, name: 'Ann' },
      { born: null, name: 'Bob' },
    ]);
    expect(Object.keys(rows[0])).toEqual(['born', 'name']);
    expect(calls).toEqual([
      [1, 1],
      [2, 1],
    ]);
  });

  it('uses the key as the header when no header is given', () => {
    expect(sheetToObjects(sheetOf([['a'], [1]]), { columns: [{ key: 'a' }] })).toEqual([{ a: 1 }]);
  });

  it('matches the leftmost column when a header is repeated', () => {
    expect(
      sheetToObjects(
        sheetOf([
          ['a', 'a'],
          [1, 2],
        ]),
        { columns: [{ key: 'a' }] }
      )
    ).toEqual([{ a: 1 }]);
  });

  it('throws INVALID_FILE naming the missing column, unless it is optional', () => {
    const failure = failureOf(() =>
      sheetToObjects<Person>(sheet, {
        columns: [
          { key: 'name', header: 'Full name' },
          { key: 'email', header: 'E-mail' },
        ],
      })
    );
    expect(failure.code).toBe('INVALID_FILE');
    expect(failure.message).toBe('Column "E-mail" not found in row 1');

    expect(
      sheetToObjects<Person>(sheet, {
        columns: [
          { key: 'name', header: 'Full name' },
          { key: 'email', header: 'E-mail', optional: true },
        ],
      })
    ).toEqual([
      { name: 'Ann', email: null },
      { name: 'Bob', email: null },
    ]);
  });

  it('throws for a missing column when the sheet has no rows at all', () => {
    const failure = failureOf(() => sheetToObjects({ data: [] }, { columns: [{ key: 'a' }] }));
    expect(failure.code).toBe('INVALID_FILE');
  });

  it('lets a column index win over the header', () => {
    expect(
      sheetToObjects(sheet, { columns: [{ key: 'who', header: 'nothing', column: 0 }] })
    ).toEqual([{ who: 'Ann' }, { who: 'Bob' }]);
  });

  it('decides which rows are empty before parse, from the raw cells', () => {
    const rows = sheetToObjects(sheetOf([['a'], [null]]), {
      columns: [{ key: 'a', parse: () => 'default' }],
    });
    expect(rows).toEqual([]);
  });
});

describe('writing objects and reading them back', () => {
  it('round-trips strings, numbers, booleans, dates and errors', () => {
    const rows = [
      {
        s: 'x',
        n: 1.5,
        b: false,
        d: new Date(2024, 1, 29),
        e: { error: '#DIV/0!' as const },
        z: null,
      },
      {
        s: 'y',
        n: -2,
        b: true,
        d: new Date(2023, 11, 31, 10, 30),
        e: { error: '#N/A' as const },
        z: null,
      },
    ];
    const columns = (Object.keys(rows[0]) as (keyof (typeof rows)[number] & string)[]).map(key => ({
      key,
    }));
    const sheet = new ExcelReader().parseFromBuffer(
      new ExcelWriter().createWorkbookBuffer([objectsToSheet(rows, columns)])
    ).sheets[0];
    expect(sheetToObjects(sheet)).toEqual(rows);
  });
});
