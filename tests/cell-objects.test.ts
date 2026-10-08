import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import {
  ExcelBridge,
  ExcelWriter,
  Workbook,
  calculateColumnWidths,
  createExcelWorkbookStream,
  streamToBuffer,
} from '../src';
import type { CellValue } from '../src';
import { cellAt, part, readFirstSheet } from './helpers/read';
import { DATE_STYLES, buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();
const sheetXml = (data: CellValue[][], sharedStrings = false) =>
  part(
    new ExcelWriter({ sharedStrings }).createWorkbookBuffer([{ data }]),
    'xl/worksheets/sheet1.xml'
  );
const cellXml = (data: CellValue[][], ref: string) =>
  new RegExp(`<c r="${ref}"[^>]*?(?:/>|>.*?</c>)`).exec(sheetXml(data))?.[0];
const written = (value: CellValue) => cellXml([[value]], 'A1');
const streamed = async (rows: CellValue[][]) =>
  streamToBuffer(createExcelWorkbookStream([{ rows }]));

describe('formula cells', () => {
  it('writes the cached result after the formula', () => {
    expect(written({ formula: 'A1*2', result: 4 })).toBe('<c r="A1"><f>A1*2</f><v>4</v></c>');
  });

  it.each([
    ['text', 'x2', '<c r="A1" t="str"><f>F</f><v>x2</v></c>'],
    ['an empty text', '', '<c r="A1" t="str"><f>F</f><v></v></c>'],
    ['a boolean', false, '<c r="A1" t="b"><f>F</f><v>0</v></c>'],
    ['an error', { error: '#N/A' as const }, '<c r="A1" t="e"><f>F</f><v>#N/A</v></c>'],
    ['zero', 0, '<c r="A1"><f>F</f><v>0</v></c>'],
  ])('writes %s as the result', (_label, result, expected) => {
    expect(written({ formula: 'F', result })).toBe(expected);
  });

  it('writes a date result with the date style', () => {
    expect(written({ formula: 'TODAY()', result: new Date(2024, 0, 15) })).toBe(
      '<c r="A1" s="1"><f>TODAY()</f><v>45306</v></c>'
    );
  });

  it('writes the same bytes as the string shorthand when there is no result', () => {
    expect(sheetXml([[{ formula: 'A1*2' }, { formula: 'B1', result: undefined }]])).toBe(
      sheetXml([['=A1*2', '=B1']])
    );
    expect(written({ formula: 'F', result: null as unknown as undefined })).toBe(
      '<c r="A1"><f>F</f></c>'
    );
  });

  it('drops one leading = from the formula', () => {
    expect(written({ formula: '=SUM(B1:B2)' })).toBe(written('=SUM(B1:B2)'));
  });

  it('escapes the formula and a text result and strips control characters', () => {
    expect(written({ formula: 'A1&"<x>"\x01', result: 'a & <b>\x0b' })).toBe(
      '<c r="A1" t="str"><f>A1&amp;"&lt;x&gt;"</f><v>a &amp; &lt;b&gt;</v></c>'
    );
  });

  it('keeps asking Excel to recalculate on open', () => {
    const bytes = writer.createWorkbookBuffer([
      { data: [[{ formula: 'A2', result: 1 }]], options: { name: 'S' } },
    ]);
    expect(part(bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });

  it.each([NaN, Infinity, -Infinity])('rejects the result %s and names the cell', result => {
    expect(() => sheetXml([['a', { formula: 'F', result }]])).toThrow(
      new RegExp(`Cell B1 holds ${result}`)
    );
  });

  it('rejects an invalid Date result and a result over the cell length limit', () => {
    expect(() => sheetXml([[{ formula: 'F', result: new Date('x') }]])).toThrow(
      'Cell A1 holds an invalid Date'
    );
    expect(() => sheetXml([[{ formula: 'F', result: 'x'.repeat(32768) }]])).toThrow(
      /exceeds Excel limit/
    );
  });

  it.each([{ formula: undefined }, { formula: 5 }, { formula: null }])(
    'rejects %j as a formula and names the cell',
    value => {
      expect(() => sheetXml([['a', value as unknown as CellValue]])).toThrow(
        'Cell B1 needs a string formula or text'
      );
    }
  );
});

describe('text cells', () => {
  it('writes text that starts with = as text', () => {
    expect(written({ text: '=SUM(A1)' })).toBe(
      '<c r="A1" t="inlineStr"><is><t>=SUM(A1)</t></is></c>'
    );
  });

  it('writes the same bytes as a plain string for other text', () => {
    const values = ['', 'abc', '  padded ', 'a & <b>', '00123'];
    expect(sheetXml([values.map(text => ({ text }))])).toBe(sheetXml([values]));
  });

  it('shares text with the strings table and with equal plain strings', () => {
    const bytes = new ExcelWriter({ sharedStrings: true }).createWorkbookBuffer([
      { data: [[{ text: '=A1' }, '=A1x', 'dup', { text: 'dup' }]] },
    ]);
    expect(part(bytes, 'xl/sharedStrings.xml')).toContain('uniqueCount="2"');
    expect(part(bytes, 'xl/sharedStrings.xml')).toContain('<si><t>=A1</t></si>');
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="A1" t="s"><v>0</v></c><c r="B1"><f>A1x</f></c><c r="C1" t="s"><v>1</v></c><c r="D1" t="s"><v>1</v></c>'
    );
  });

  it('reads back as a string without a formula', () => {
    const cell = cellAt(
      readFirstSheet(writer.createWorkbookBuffer([{ data: [[{ text: '=A1' }]] }])),
      'A1'
    );
    expect(cell).toMatchObject({ type: 'string', value: '=A1' });
    expect(cell?.formula).toBeUndefined();
  });

  it('applies the cell length limit and strips control characters like a plain string', () => {
    expect(() => sheetXml([[{ text: 'x'.repeat(32768) }]])).toThrow(/exceeds Excel limit/);
    expect(written({ text: 'a\x01b' })).toBe(written('a\x01b'));
  });

  it.each([{ text: undefined }, { text: 5 }, { text: null }])(
    'rejects %j as text and names the cell',
    value => {
      expect(() => sheetXml([['a', value as unknown as CellValue]])).toThrow(
        'Cell B1 needs a string formula or text'
      );
    }
  );
});

describe('error cells', () => {
  const errors = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'] as const;

  it.each(errors)('writes %s and reads it back as an error', error => {
    expect(written({ error })).toBe(`<c r="A1" t="e"><v>${error}</v></c>`);
    const bytes = writer.createWorkbookBuffer([{ data: [[{ error }]] }]);
    expect(cellAt(readFirstSheet(bytes), 'A1')).toMatchObject({ type: 'error', value: error });
  });

  it.each(['#SPILL!', '#n/a', 'N/A', '', '#DIV/0', '#N/A\n'])('rejects %j', error => {
    expect(() => sheetXml([['a', { error: error as '#N/A' }]])).toThrow(
      `Cell B1 holds ${error}, which a worksheet cannot store`
    );
  });

  it('rejects an error that is not a string', () => {
    expect(() => sheetXml([[{ error: undefined as unknown as '#N/A' }]])).toThrow(
      /Cell A1 holds undefined/
    );
  });
});

describe('values that keep their behaviour', () => {
  it('still writes a string that starts with = as a formula', () => {
    expect(written('=1+1')).toBe('<c r="A1"><f>1+1</f></c>');
  });

  it('still writes an object without a cell key as its string form', () => {
    class Price {
      toString() {
        return '9.99 USD';
      }
    }
    expect(written(new Price() as unknown as CellValue)).toBe(
      '<c r="A1" t="inlineStr"><is><t>9.99 USD</t></is></c>'
    );
  });

  it('still writes null and undefined as empty cells', () => {
    expect(sheetXml([[null, undefined]])).toContain('<c r="A1"/><c r="B1"/>');
  });
});

describe('column widths', () => {
  it('measures a text cell like the plain string', () => {
    expect(calculateColumnWidths([[{ text: 'x'.repeat(30) }]])).toEqual(
      calculateColumnWidths([['x'.repeat(30)]])
    );
  });

  it('measures the result of a formula and the name of an error', () => {
    expect(calculateColumnWidths([[{ formula: 'A1', result: 'x'.repeat(30) }]])).toEqual(
      calculateColumnWidths([['x'.repeat(30)]])
    );
    expect(calculateColumnWidths([[{ error: '#DIV/0!' }]])).toEqual(
      calculateColumnWidths([['#DIV/0!']])
    );
  });

  it('measures a formula without a result like the string shorthand', () => {
    expect(calculateColumnWidths([[{ formula: 'B2*C2*D2*E2*F2*G2*H2*I2' }]])).toEqual(
      calculateColumnWidths([['=B2*C2*D2*E2*F2*G2*H2*I2']])
    );
  });

  it('does not change the widths of plain values', () => {
    const row = ['a', 1, true, null, undefined, '=B2*C2*D2*E2*F2*G2', new Date(2024, 0, 1), 'W@M'];
    expect(calculateColumnWidths([row])).toEqual([8, 8, 8, 8, 8, 24, 50, 8]);
  });
});

describe('streaming writer', () => {
  const rows: CellValue[][] = [
    ['Id', { formula: 'A2*2', result: 4 }, { text: '=raw' }, { error: '#REF!' }],
    [
      { formula: 'TODAY()', result: new Date(2024, 0, 15) },
      { formula: 'A1', result: 'x' },
    ],
  ];

  it('writes the same sheet cells as ExcelWriter', async () => {
    const buffer = await streamed(rows);
    const fromWriter = writer.createWorkbookBuffer([{ data: rows }]);
    const body = (bytes: Uint8Array) =>
      /<sheetData>[\s\S]*<\/sheetData>/.exec(part(bytes, 'xl/worksheets/sheet1.xml'))?.[0];
    expect(body(buffer)).toBe(body(fromWriter));
    expect(part(buffer, 'xl/styles.xml')).toBe(part(fromWriter, 'xl/styles.xml'));
  });

  it('rejects an invalid error with the cell, after the earlier rows were produced', async () => {
    await expect(streamed([['a'], ['b', { error: '#BAD' as '#N/A' }]])).rejects.toThrow(
      'Cell B2 holds #BAD, which a worksheet cannot store'
    );
  });

  it('reads back with the library reader', async () => {
    const sheet = ExcelBridge.read(await streamed(rows)).sheets[0];
    expect(cellAt(sheet, 'B1')).toMatchObject({ type: 'number', value: 4, formula: 'A2*2' });
    expect(cellAt(sheet, 'C1')).toMatchObject({ type: 'string', value: '=raw' });
    expect(cellAt(sheet, 'D1')).toMatchObject({ type: 'error', value: '#REF!' });
    expect(cellAt(sheet, 'A2')).toMatchObject({ type: 'date', formula: 'TODAY()' });
  });
});

describe('reader', () => {
  const body = (cells: string) =>
    buildXlsx(`<sheetData><row r="1">${cells}</row></sheetData>`, {}, { styles: DATE_STYLES });

  it('returns the cached value, its type and the formula of a formula cell', () => {
    const sheet = readFirstSheet(
      body(
        [
          '<c r="A1"><f>1+1</f><v>2</v></c>',
          '<c r="B1" t="str"><f>"0"&amp;"7"</f><v>07</v></c>',
          '<c r="C1" t="b"><f>TRUE()</f><v>1</v></c>',
          '<c r="D1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
          '<c r="E1" s="1"><f>TODAY()</f><v>45306</v></c>',
          '<c r="F1"><f>NOW()</f></c>',
        ].join('')
      )
    );
    expect(
      sheet.data[0].map(({ type, value, formula }) => ({ type, value, formula }))
    ).toMatchObject([
      { type: 'number', value: 2, formula: '1+1' },
      { type: 'string', value: '07', formula: '"0"&"7"' },
      { type: 'boolean', value: true, formula: 'TRUE()' },
      { type: 'error', value: '#DIV/0!', formula: '1/0' },
      { type: 'date', formula: 'TODAY()' },
      { type: 'empty', value: null, formula: 'NOW()' },
    ]);
  });

  it('returns text that starts with = without a formula', () => {
    const cell = cellAt(
      readFirstSheet(body('<c r="A1" t="inlineStr"><is><t>=SUM(A1)</t></is></c>')),
      'A1'
    );
    expect(cell).toMatchObject({ type: 'string', value: '=SUM(A1)' });
    expect('formula' in (cell ?? {})).toBe(false);
  });
});

describe('Workbook', () => {
  it('writes the cell objects it is given and returns them', () => {
    const workbook = Workbook.create();
    workbook.addSheet('S', [['a']]);
    workbook.setCellValue('S', 0, 1, { formula: 'A1', result: 'a' });
    workbook.setCellValue('S', 0, 2, { error: '#N/A' });
    workbook.setCellValue('S', 0, 3, { text: '=x' });
    expect(workbook.getCellValue('S', 0, 3)).toEqual({ text: '=x' });
    const sheet = readFirstSheet(workbook.toBuffer());
    expect(cellAt(sheet, 'B1')).toMatchObject({ type: 'string', value: 'a', formula: 'A1' });
    expect(cellAt(sheet, 'C1')).toMatchObject({ type: 'error', value: '#N/A' });
    expect(cellAt(sheet, 'D1')).toMatchObject({ type: 'string', value: '=x' });
  });

  it('keeps text that starts with = as text through a save', () => {
    const original = writer.createWorkbookBuffer([{ data: [[{ text: '=SUM(A1)' }, '=1+1']] }]);
    const workbook = Workbook.fromBuffer(original);
    expect(workbook.getCellValue('Sheet1', 0, 0)).toBe('=SUM(A1)');
    expect(workbook.getCellValue('Sheet1', 0, 1)).toBe('=1+1');
    expect(part(workbook.toBuffer(), 'xl/worksheets/sheet1.xml')).toBe(
      part(original, 'xl/worksheets/sheet1.xml')
    );
  });

  it('saves a formula without its cached result, which it cannot recalculate', () => {
    const original = writer.createWorkbookBuffer([{ data: [[2, { formula: 'A1*2', result: 4 }]] }]);
    const workbook = Workbook.fromBuffer(original);
    expect(workbook.getCellValue('Sheet1', 0, 1)).toBe('=A1*2');
    expect(cellAt(readFirstSheet(workbook.toBuffer()), 'B1')).toMatchObject({
      type: 'empty',
      formula: 'A1*2',
    });
  });

  it('does not turn text from a loaded file into a formula', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>=HYPERLINK("https://example.com","x")</t></is></c></row></sheetData>'
    );
    const saved = Workbook.fromBuffer(bytes).toBuffer();
    expect(part(saved, 'xl/worksheets/sheet1.xml')).not.toContain('<f>');
    expect(cellAt(readFirstSheet(saved), 'A1')).toMatchObject({
      type: 'string',
      value: '=HYPERLINK("https://example.com","x")',
    });
  });

  it('keeps the error cells it loads', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row></sheetData>'
    );
    expect(cellAt(readFirstSheet(Workbook.fromBuffer(bytes).toBuffer()), 'A1')).toMatchObject({
      type: 'error',
      value: '#DIV/0!',
    });
  });

  it('saves error text outside the seven classic errors as text', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" t="e"><v>#SPILL!</v></c></row></sheetData>'
    );
    expect(cellAt(readFirstSheet(Workbook.fromBuffer(bytes).toBuffer()), 'A1')).toMatchObject({
      type: 'string',
      value: '#SPILL!',
    });
  });
});

describe('files other tools can read', () => {
  it('lists the zip parts unchanged', () => {
    const bytes = writer.createWorkbookBuffer([{ data: [[{ formula: 'A1', result: 1 }]] }]);
    expect(Object.keys(unzipSync(bytes)).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'docProps/app.xml',
      'docProps/core.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
  });
});
