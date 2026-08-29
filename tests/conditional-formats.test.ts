import { describe, it, expect } from 'vitest';
import { ExcelWriter, ExcelBridge, Workbook } from '../src';
import type { ConditionalFormat } from '../src';

const sheet = () => ({
  data: [
    ['Item', 'Qty'],
    ['A', 5],
    ['B', 20],
  ] as (string | number)[][],
});

describe('Conditional formatting round-trip', () => {
  it('reads back a cellValue rule with its style', () => {
    const formats: ConditionalFormat[] = [
      {
        type: 'cellValue',
        range: 'B2:B3',
        operator: 'lessThan',
        value: 10,
        style: { background: '#FFC7CE', color: '#9C0006', bold: true },
      },
    ];

    const buffer = new ExcelWriter().createWorkbookBuffer([{ ...sheet(), conditionalFormats: formats }]);
    const parsed = ExcelBridge.read(buffer);

    expect(parsed.sheets[0].conditionalFormats).toEqual(formats);
  });

  it('reads back a between rule with value2', () => {
    const formats: ConditionalFormat[] = [
      {
        type: 'cellValue',
        range: 'B2:B3',
        operator: 'between',
        value: 1,
        value2: 100,
        style: { italic: true, color: '#006100' },
      },
    ];

    const buffer = new ExcelWriter().createWorkbookBuffer([{ ...sheet(), conditionalFormats: formats }]);
    const parsed = ExcelBridge.read(buffer);

    expect(parsed.sheets[0].conditionalFormats).toEqual(formats);
  });

  it('reads back an expression rule', () => {
    const formats: ConditionalFormat[] = [
      {
        type: 'expression',
        range: 'A2:B3',
        formula: '$B2<10',
        style: { background: '#FFE6E6', bold: true },
      },
    ];

    const buffer = new ExcelWriter().createWorkbookBuffer([{ ...sheet(), conditionalFormats: formats }]);
    const parsed = ExcelBridge.read(buffer);

    expect(parsed.sheets[0].conditionalFormats).toEqual(formats);
  });

  it('reads back a two- and three-color scale', () => {
    const formats: ConditionalFormat[] = [
      { type: 'colorScale', range: 'B2:B3', colors: ['#F8696B', '#63BE7B'] },
      { type: 'colorScale', range: 'B2:B3', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
    ];

    const buffer = new ExcelWriter().createWorkbookBuffer([{ ...sheet(), conditionalFormats: formats }]);
    const parsed = ExcelBridge.read(buffer);

    expect(parsed.sheets[0].conditionalFormats).toEqual(formats);
  });

  it('preserves conditional formats through a Workbook load/edit/save cycle', () => {
    const formats: ConditionalFormat[] = [
      {
        type: 'cellValue',
        range: 'B2:B3',
        operator: 'greaterThan',
        value: 15,
        style: { background: '#E2EFDA' },
      },
      { type: 'colorScale', range: 'B2:B3', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
    ];

    const original = new ExcelWriter().createWorkbookBuffer([{ ...sheet(), conditionalFormats: formats }]);

    const workbook = Workbook.fromBuffer(original);
    const resaved = workbook.toBuffer();
    const parsed = ExcelBridge.read(resaved);

    expect(parsed.sheets[0].conditionalFormats).toEqual(formats);
  });

  it('omits the field when a sheet has no conditional formats', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([sheet()]);
    const parsed = ExcelBridge.read(buffer);

    expect(parsed.sheets[0].conditionalFormats).toBeUndefined();
  });
});
