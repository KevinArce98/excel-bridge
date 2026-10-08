import { describe, it, expect } from 'vitest';
import {
  ExcelReader,
  ExcelWriter,
  createExcelWorkbookStream,
  objectsToSheet,
  objectsToStreamingSheet,
  streamToBuffer,
} from '../src';
import type { WriteColumn } from '../src';
import { cellAt, part } from './helpers/read';

describe('objectsToSheet with keys that exist on every object', () => {
  it.each(['constructor', 'toString', 'valueOf', 'hasOwnProperty', '__proto__'])(
    'writes an empty cell for a missing %s',
    key => {
      const sheet = objectsToSheet(
        [{ id: 1 } as Record<string, unknown>],
        [{ key: 'id' }, { key }]
      );
      expect(sheet.data[1]).toEqual([1, undefined]);
    }
  );

  it('still reads a property the row owns under one of those names', () => {
    const sheet = objectsToSheet(
      [{ toString: 'own' } as Record<string, unknown>],
      [{ key: 'toString' }]
    );
    expect(sheet.data[1]).toEqual(['own']);
  });

  it('still reads a getter defined on the row class', () => {
    class Line {
      get total() {
        return 42;
      }
    }
    const sheet = objectsToSheet([new Line()], [{ key: 'total' }]);
    expect(sheet.data[1]).toEqual([42]);
  });
});

interface Order {
  id: number;
  customer: string;
  total: number;
  placed: Date;
  note?: string | null;
}

const orders: Order[] = [
  { id: 1, customer: 'Ann', total: 1250.5, placed: new Date(2024, 0, 15) },
  { id: 2, customer: 'Bob', total: 80, placed: new Date(2024, 1, 2), note: 'rush' },
];

const read = (bytes: Uint8Array) => new ExcelReader().parseFromBuffer(bytes).sheets[0];

describe('objectsToSheet', () => {
  it('writes a header row from header or key, then one row per object', () => {
    const sheet = objectsToSheet<Order>(orders, [
      { key: 'id' },
      { key: 'customer', header: 'Customer' },
      { key: 'note' },
    ]);

    expect(sheet.data).toEqual([
      ['id', 'Customer', 'note'],
      [1, 'Ann', undefined],
      [2, 'Bob', 'rush'],
    ]);
    expect(sheet.styles).toBeUndefined();
    expect(sheet.options).toEqual({});
  });

  it('writes only the header for no rows, and accepts any iterable', () => {
    expect(objectsToSheet<Order>([], [{ key: 'id' }]).data).toEqual([['id']]);
    function* generated() {
      yield* orders;
    }
    expect(objectsToSheet(generated(), [{ key: 'id' }]).data).toEqual([['id'], [1], [2]]);
  });

  it('ignores properties without a column and leaves a missing property empty', () => {
    const { data } = objectsToSheet<Record<string, unknown>>(
      [{ a: 1, extra: 'x' }, { b: 2 }],
      [{ key: 'a' }, { key: 'b' }]
    );
    expect(data).toEqual([
      ['a', 'b'],
      [1, undefined],
      [undefined, 2],
    ]);
  });

  it('styles the header cells with headerStyle and the body cells with the column style', () => {
    const header = { bold: true, background: '#4472C4' };
    const sheet = objectsToSheet<Order>(
      orders,
      [
        { key: 'id' },
        { key: 'total', style: { bold: true, numberFormat: '0' }, numberFormat: '#,##0.00' },
        { key: 'placed', numberFormat: 'yyyy-mm-dd' },
      ],
      { headerStyle: header }
    );

    expect(sheet.styles).toEqual({
      '0-0': header,
      '0-1': header,
      '0-2': header,
      '1-1': { bold: true, numberFormat: '#,##0.00' },
      '2-1': { bold: true, numberFormat: '#,##0.00' },
      '1-2': { numberFormat: 'yyyy-mm-dd' },
      '2-2': { numberFormat: 'yyyy-mm-dd' },
    });
  });

  it('turns the widths into columnWidths with holes for columns without one', () => {
    const { options } = objectsToSheet<Order>(orders, [
      { key: 'id' },
      { key: 'customer', width: 24 },
      { key: 'total', width: 12 },
    ]);
    expect(options?.columnWidths).toHaveLength(3);
    expect(0 in options!.columnWidths!).toBe(false);
    expect(options?.columnWidths?.[1]).toBe(24);
  });

  it('passes the sheet options through and keeps their columnWidths when no column has a width', () => {
    const { options } = objectsToSheet<Order>(orders, [{ key: 'id' }], {
      name: 'Orders',
      freezePane: { row: 1 },
      columnWidths: [30],
      autoFilter: { range: 'A1:A3' },
    });
    expect(options).toEqual({
      name: 'Orders',
      freezePane: { row: 1 },
      columnWidths: [30],
      autoFilter: { range: 'A1:A3' },
    });
  });

  it('lets a column width replace options.columnWidths', () => {
    const { options } = objectsToSheet<Order>(orders, [{ key: 'id', width: 5 }], {
      columnWidths: [30],
      autoWidth: true,
    });
    expect(options?.columnWidths).toEqual([5]);
  });

  it('produces a file that Excel-style readers see as written', () => {
    const sheet = objectsToSheet<Order>(
      orders,
      [
        { key: 'id', header: 'Id', width: 6 },
        { key: 'total', header: 'Total', width: 14, numberFormat: '#,##0.00' },
        { key: 'placed', header: 'Placed', width: 14, numberFormat: 'yyyy-mm-dd' },
      ],
      { name: 'Orders', headerStyle: { bold: true } }
    );
    const bytes = new ExcelWriter().createWorkbookBuffer([sheet]);
    const parsed = read(bytes);

    expect(parsed.name).toBe('Orders');
    expect(parsed.columnWidths).toEqual([6, 14, 14]);
    expect(parsed.data[1].map(cell => cell.value)).toEqual([1, 1250.5, new Date(2024, 0, 15)]);
    expect(parsed.styles?.['0-0']).toMatchObject({ bold: true });
    expect(parsed.styles?.['1-1']).toMatchObject({ numberFormat: '#,##0.00' });
    expect(parsed.styles?.['1-2']).toMatchObject({ numberFormat: 'yyyy-mm-dd' });
  });

  it('writes a string that starts with = as text, not as a formula', () => {
    const sheet = objectsToSheet([{ a: '=1+1' }], [{ key: 'a' }]);
    const cell = cellAt(read(new ExcelWriter().createWorkbookBuffer([sheet])), 'A2');

    expect(cell).toMatchObject({ type: 'string', value: '=1+1' });
    expect(cell).not.toHaveProperty('formula');
  });

  it('checks the column keys against the row type', () => {
    // @ts-expect-error 'custmer' is not a key of Order
    const columns: WriteColumn<Order>[] = [{ key: 'custmer' }];
    expect(columns).toHaveLength(1);
  });
});

describe('objectsToStreamingSheet', () => {
  async function* fromDatabase() {
    yield* orders;
  }

  it('streams a header and the objects of an async iterable', async () => {
    const input = objectsToStreamingSheet<Order>(
      fromDatabase(),
      [
        { key: 'id', width: 6 },
        { key: 'customer', header: 'Customer', width: 20 },
      ],
      { name: 'Orders', freezePane: { row: 1 }, headerStyle: { bold: true } }
    );
    const bytes = await streamToBuffer(createExcelWorkbookStream([input]));
    const sheet = read(bytes);

    expect(sheet.name).toBe('Orders');
    expect(sheet.freezePane).toEqual({ row: 1 });
    expect(sheet.columnWidths).toEqual([6, 20]);
    expect(sheet.data.map(row => row.map(cell => cell.value))).toEqual([
      ['id', 'Customer'],
      [1, 'Ann'],
      [2, 'Bob'],
    ]);
    expect(sheet.styles?.['0-1']).toMatchObject({ bold: true });
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<col min="1" max="1" width="6"');
  });

  it('can be consumed twice when the source can', async () => {
    const input = objectsToStreamingSheet<Order>(orders, [{ key: 'id' }]);
    const once = read(await streamToBuffer(createExcelWorkbookStream([input])));
    const twice = read(await streamToBuffer(createExcelWorkbookStream([input])));
    expect(twice.data).toEqual(once.data);
    expect(once.data).toHaveLength(3);
  });

  it('merges explicit styles over the header styles', async () => {
    const input = objectsToStreamingSheet<Order>(orders, [{ key: 'id' }, { key: 'total' }], {
      headerStyle: { bold: true },
      styles: { '0-1': { italic: true }, '1-1': { background: '#FFFF00' } },
    });
    const sheet = read(await streamToBuffer(createExcelWorkbookStream([input])));
    expect(sheet.styles?.['0-0']).toMatchObject({ bold: true });
    expect(sheet.styles?.['0-1']).toMatchObject({ italic: true });
    expect(sheet.styles?.['1-1']).toMatchObject({ background: '#FFFF00' });
  });

  it('rejects a column style, which a streaming sheet cannot apply to every row', () => {
    expect(() =>
      objectsToStreamingSheet<Order>(orders, [{ key: 'total', numberFormat: '0.00' } as never])
    ).toThrow('Streaming column "total" cannot have a style');
    expect(() =>
      objectsToStreamingSheet<Order>(orders, [{ key: 'total', style: { bold: true } } as never])
    ).toThrow(expect.objectContaining({ code: 'INVALID_INPUT' }));
  });

  it('does not start reading the source before the stream is consumed', () => {
    let pulled = false;
    async function* spy() {
      pulled = true;
      yield orders[0];
    }
    objectsToStreamingSheet(spy(), [{ key: 'id' }]);
    expect(pulled).toBe(false);
  });
});
