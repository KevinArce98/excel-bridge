import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { ExcelWriter, createExcelWorkbookStream, streamToBuffer, dataValidation } from '../src';
import type { ExcelData } from '../src';

const parts = (buffer: Uint8Array): Record<string, string> => {
  const files = unzipSync(buffer);
  return Object.fromEntries(
    Object.keys(files)
      .filter(name => name !== 'docProps/core.xml')
      .sort()
      .map(name => [name, strFromU8(files[name])])
  );
};

const kitchenSink = (): ExcelData[] => [
  {
    data: [
      ['Product', 'Price', 'Qty', 'Total', 'Due'],
      ['Laptop', 999.99, 5, '=B2*C2', new Date(2024, 0, 15)],
      ['Mouse', 29.99, 20, '=B3*C3', new Date(2024, 1, 20, 9, 30)],
      [' padded ', true, null, 'x & <y> "z"', undefined],
    ],
    styles: {
      '0-0': { background: '#4472C4', bold: true, color: '#FFFFFF', border: true },
      '1-1': { numberFormat: '#,##0.00', underline: true, italic: true },
      '2-0': { align: 'center', verticalAlign: 'middle', wrapText: true, fontSize: 14 },
      '3-0': { fontName: 'Arial', color: '#abc' },
    },
    mergeCells: ['A4:B4'],
    validations: [
      dataValidation.list('C2:C3', ['1', '2']),
      dataValidation.wholeNumber('C2:C3', 'between', 1, 5),
      dataValidation.decimal('B2:B3', 'greaterThanOrEqual', 0),
      dataValidation.textLength('A2:A3', 'lessThan', 20),
      dataValidation.dateBetween('E2:E3', new Date(2024, 0, 1), new Date(2024, 11, 31)),
    ],
    conditionalFormats: [
      {
        type: 'cellValue',
        range: 'C2:C3',
        operator: 'between',
        value: 1,
        value2: 10,
        style: { background: '#FFC7CE', color: '#9C0006' },
      },
      { type: 'expression', range: 'A2:D3', formula: '$C2>5', style: { bold: true } },
      { type: 'colorScale', range: 'B2:B3', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
    ],
    options: { name: 'Sales & Co', freezePane: { row: 1, col: 1 }, columnWidths: [20, 10] },
  },
  {
    data: [
      ['Name', 'Notes'],
      ['A', 'a longer value'],
    ],
    options: { autoWidth: true, freezePane: {} },
  },
];

describe('Output snapshots', () => {
  it('ExcelWriter parts stay stable', () => {
    const buffer = new ExcelWriter({
      creator: 'Snapshot',
      title: 'T',
      subject: 'S',
    }).createWorkbookBuffer(kitchenSink());
    expect(parts(buffer)).toMatchSnapshot();
  });

  it('ExcelWriter parts with shared strings stay stable', () => {
    const buffer = new ExcelWriter({ sharedStrings: true }).createWorkbookBuffer(kitchenSink());
    expect(parts(buffer)).toMatchSnapshot();
  });

  it('streaming writer parts stay stable', async () => {
    function* rows() {
      yield ['Id', 'Name', 'Joined'];
      for (let i = 1; i <= 3; i++) yield [i, `Row ${i}`, new Date(2024, 0, i)];
    }
    const buffer = await streamToBuffer(
      createExcelWorkbookStream(
        [
          {
            name: 'Data',
            rows: rows(),
            styles: { '0-0': { bold: true, background: '#DDEBF7' } },
            freezePane: { row: 1 },
            columnWidths: [8, 20],
            mergeCells: ['B4:C4'],
          },
          { rows: [[1, 2]], freezePane: { row: 0, col: 0 } },
        ],
        { creator: 'Stream' }
      )
    );
    expect(parts(buffer)).toMatchSnapshot();
  });
});
