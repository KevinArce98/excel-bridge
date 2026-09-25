import { describe, it, expect } from 'vitest';
import { ExcelWriter, dataValidation, createExcelWorkbookStream, streamToBuffer } from '../src';
import type { ConditionalFormat } from '../src';
import { unzipSync, strFromU8 } from 'fflate';

const parts = (buffer: Uint8Array) => {
  const files = unzipSync(buffer);
  return {
    workbook: strFromU8(files['xl/workbook.xml']),
    sheet1: strFromU8(files['xl/worksheets/sheet1.xml']),
  };
};

const expectOrdered = (xml: string, markers: string[]) => {
  const positions = markers.map(marker => xml.indexOf(marker));
  markers.forEach((marker, index) => expect(positions[index], marker).toBeGreaterThan(-1));
  expect(positions).toEqual([...positions].sort((a, b) => a - b));
};

const everySection = {
  data: [
    ['A', 'B'],
    ['C', 'D'],
  ],
  mergeCells: ['A2:B2'],
  conditionalFormats: [
    { type: 'colorScale', range: 'A1:A2', colors: ['#F8696B', '#63BE7B'] },
  ] as ConditionalFormat[],
  validations: [dataValidation.list('B2:B2', ['x', 'y'])],
  hyperlinks: [{ range: 'B1', url: 'https://example.com' }],
  options: { freezePane: { row: 1 }, columnWidths: [10, 12], autoFilter: { range: 'A1:B2' } },
};

describe('OOXML element ordering (Excel schema compliance)', () => {
  it('workbook.xml places <sheets> before <calcPr>', () => {
    const { workbook } = parts(new ExcelWriter().createWorkbookBuffer([{ data: [['A']] }]));
    expect(workbook.indexOf('<sheets>')).toBeGreaterThan(-1);
    expect(workbook.indexOf('<calcPr')).toBeGreaterThan(-1);
    expect(workbook.indexOf('<sheets>')).toBeLessThan(workbook.indexOf('<calcPr'));
  });

  it('worksheet orders mergeCells → conditionalFormatting → dataValidations', () => {
    const conditionalFormats: ConditionalFormat[] = [
      { type: 'colorScale', range: 'A1:A2', colors: ['#F8696B', '#63BE7B'] },
    ];
    const buffer = new ExcelWriter().createWorkbookBuffer([
      {
        data: [
          ['A', 'B'],
          ['C', 'D'],
        ],
        mergeCells: ['A1:B1'],
        conditionalFormats,
        validations: [dataValidation.list('A2:A2', ['x', 'y'])],
      },
    ]);
    const { sheet1 } = parts(buffer);

    const merge = sheet1.indexOf('<mergeCells');
    const cf = sheet1.indexOf('<conditionalFormatting');
    const validation = sheet1.indexOf('<dataValidations');

    expect(merge).toBeGreaterThan(-1);
    expect(cf).toBeGreaterThan(-1);
    expect(validation).toBeGreaterThan(-1);
    expect(merge).toBeLessThan(cf);
    expect(cf).toBeLessThan(validation);
  });

  it('worksheet orders every section per the schema', () => {
    const { sheet1 } = parts(new ExcelWriter().createWorkbookBuffer([everySection]));

    expectOrdered(sheet1, [
      '<sheetViews',
      '<cols',
      '<sheetData',
      '</sheetData>',
      '<autoFilter',
      '<mergeCells',
      '<conditionalFormatting',
      '<dataValidations',
      '<hyperlinks',
      '</worksheet>',
    ]);
  });

  it('workbook places definedNames between sheets and calcPr', () => {
    const { workbook } = parts(new ExcelWriter().createWorkbookBuffer([everySection]));

    expectOrdered(workbook, ['</sheets>', '<definedNames>', '<calcPr']);
  });

  it('streaming worksheet follows the same order', async () => {
    const { data, mergeCells, hyperlinks, options } = everySection;
    const buffer = await streamToBuffer(
      createExcelWorkbookStream([{ rows: data, mergeCells, hyperlinks, ...options }])
    );

    expectOrdered(parts(buffer).sheet1, [
      '<sheetViews',
      '<cols',
      '<sheetData',
      '</sheetData>',
      '<autoFilter',
      '<mergeCells',
      '<hyperlinks',
      '</worksheet>',
    ]);
  });

  it('keeps sheetData before trailing sections', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: [['A']], mergeCells: ['A1:A1'] },
    ]);
    const { sheet1 } = parts(buffer);
    expect(sheet1.indexOf('</sheetData>')).toBeLessThan(sheet1.indexOf('<mergeCells'));
  });
});
