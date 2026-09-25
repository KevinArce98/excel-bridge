import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8, zipSync, strToU8 } from 'fflate';
import { ExcelWriter, ExcelBridge, Workbook } from '../src';

const numberLikeTexts = ['00123', '1e3', '1.50', '-0', '0x1F', ' 007 ', 'true', '+5'];

const replacePart = (buffer: Uint8Array, path: string, replace: (xml: string) => string) => {
  const files = unzipSync(buffer);
  files[path] = strToU8(replace(strFromU8(files[path])));
  return zipSync(files);
};

describe('Text fidelity on read', () => {
  it('keeps number-like inline strings exact', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([{ data: [numberLikeTexts] }]);
    const row = ExcelBridge.read(buffer).sheets[0].data[0];

    expect(row.map(cell => cell.value)).toEqual(numberLikeTexts);
    expect(row.every(cell => cell.type === 'string')).toBe(true);
  });

  it('keeps number-like shared strings exact', () => {
    const buffer = new ExcelWriter({ sharedStrings: true }).createWorkbookBuffer([
      { data: [numberLikeTexts] },
    ]);

    expect(ExcelBridge.read(buffer).sheets[0].data[0].map(cell => cell.value)).toEqual(
      numberLikeTexts
    );
  });

  it('keeps rich-text runs exact in shared and inline strings', () => {
    const base = new ExcelWriter({ sharedStrings: true }).createWorkbookBuffer([
      { data: [['shared', 'inline']] },
    ]);
    const withSharedRuns = replacePart(base, 'xl/sharedStrings.xml', xml =>
      xml.replace(
        '<si><t>shared</t></si>',
        '<si><r><t>00</t></r><r><rPr><b/></rPr><t>123</t></r></si>'
      )
    );
    const withInlineRuns = replacePart(withSharedRuns, 'xl/worksheets/sheet1.xml', xml =>
      xml.replace(
        /<c r="B1"[^>]*>.*?<\/c>/,
        '<c r="B1" t="inlineStr"><is><r><t>1e3</t></r></is></c>'
      )
    );

    expect(ExcelBridge.read(withInlineRuns).sheets[0].data[0].map(cell => cell.value)).toEqual([
      '00123',
      '1e3',
    ]);
  });

  it('keeps formula text and string results exact', () => {
    const base = new ExcelWriter().createWorkbookBuffer([{ data: [['placeholder', '=1E3']] }]);
    const withStringResult = replacePart(base, 'xl/worksheets/sheet1.xml', xml =>
      xml.replace(
        /<c r="A1"[^>]*>.*?<\/c>/,
        '<c r="A1" t="str"><f>CONCAT("00","123")</f><v>00123</v></c>'
      )
    );
    const [stringResult, formula] = ExcelBridge.read(withStringResult).sheets[0].data[0];

    expect(stringResult).toMatchObject({
      value: '00123',
      type: 'string',
      formula: 'CONCAT("00","123")',
    });
    expect(formula.formula).toBe('1E3');
  });

  it('still reads numbers, booleans and dates as typed values', () => {
    const date = new Date(2024, 0, 15);
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: [[1.5, 7, 0, -0.002, 1e21, true, false, date]] },
    ]);
    const row = ExcelBridge.read(buffer).sheets[0].data[0];

    expect(row.map(cell => cell.value)).toEqual([1.5, 7, 0, -0.002, 1e21, true, false, date]);
    expect(row.map(cell => cell.type)).toEqual([
      'number',
      'number',
      'number',
      'number',
      'number',
      'boolean',
      'boolean',
      'date',
    ]);
  });

  it('keeps number-like text through a Workbook round-trip', () => {
    const workbook = Workbook.fromBuffer(
      new ExcelWriter().createWorkbookBuffer([
        { data: [['00123', 42]], options: { name: 'Codes' } },
      ])
    );

    expect(workbook.getCellValue('Codes', 0, 0)).toBe('00123');
    expect(ExcelBridge.read(workbook.toBuffer()).sheets[0].data[0].map(cell => cell.value)).toEqual(
      ['00123', 42]
    );
  });
});
