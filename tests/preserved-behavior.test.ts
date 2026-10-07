import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync } from 'fflate';
import { ExcelBridge, ExcelWriter, Workbook, hyperlink } from '../src';
import type { ConditionalFormat } from '../src';
import type { ParsedSheet } from '../src/reader';
import { buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();

const part = (bytes: Uint8Array, path: string): string => strFromU8(unzipSync(bytes)[path]);

const cellAt = (sheet: ParsedSheet, coordinate: string) =>
  sheet.data.flat().find(cell => cell?.coordinate === coordinate);

describe('Workbook round trip keeps the features the writer produces', () => {
  const conditionalFormats: ConditionalFormat[] = [
    {
      type: 'cellValue',
      range: 'B2:B4',
      operator: 'greaterThan',
      value: 10,
      style: { bold: true },
    },
    {
      type: 'cellValue',
      range: 'B2:B4',
      operator: 'between',
      value: 1,
      value2: 5,
      style: { background: '#FFC7CE' },
    },
    { type: 'expression', range: 'A2:A4', formula: '$B2>3', style: { color: '#9C0006' } },
    { type: 'colorScale', range: 'B2:B4', colors: ['#F8696B', '#63BE7B'] },
    { type: 'colorScale', range: 'B2:B4', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  ];

  const original = writer.createWorkbookBuffer([
    {
      data: [
        ['Name', 'Qty', 'Link'],
        ['pen', 4, 'docs'],
        ['ink', 12, 'mail'],
        ['pad', 2, 'jump'],
      ],
      conditionalFormats,
      hyperlinks: [
        hyperlink.url('C2', 'https://example.com/docs', { tooltip: 'Open the docs' }),
        hyperlink.email('C3', 'team@example.com', { subject: 'Question' }),
        hyperlink.internal('C4', 'Other', 'B2'),
      ],
      mergeCells: ['A6:C6'],
      options: {
        name: 'Main',
        freezePane: { row: 1, col: 1 },
        columnWidths: [18, 8.43, 30],
        autoFilter: { range: 'A1:C4' },
      },
    },
    { data: [['x', 'y']], options: { name: 'Other' } },
  ]);
  const saved = Workbook.fromBuffer(original).toBuffer();

  it.each(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'])('writes %s unchanged', path => {
    expect(part(saved, path)).toBe(part(original, path));
  });

  it('writes the hyperlink relationships unchanged', () => {
    expect(part(saved, 'xl/worksheets/_rels/sheet1.xml.rels')).toBe(
      part(original, 'xl/worksheets/_rels/sheet1.xml.rels')
    );
  });

  it('reads the same rules, links and layout back', () => {
    const before = ExcelBridge.read(original).sheets[0];
    const after = ExcelBridge.read(saved).sheets[0];
    expect(after.conditionalFormats).toEqual(before.conditionalFormats);
    expect(after.hyperlinks).toEqual(before.hyperlinks);
    expect(after.autoFilter).toEqual({ range: 'A1:C4' });
    expect(after.mergeCells).toEqual(['A6:C6']);
    expect(after.freezePane).toEqual({ row: 1, col: 1 });
    expect(after.columnWidths).toEqual([18, 8.43, 30]);
  });
});

describe('text survives writing and reading', () => {
  const values = [
    '00123',
    '007',
    '000',
    '01234-0001',
    'Zoë Müller',
    '日本語のテキスト',
    '😀 👨‍👩‍👧 🇨🇷',
    'مرحبا بالعالم',
    '<a href="x">&amp; "quoted" \'single\'</a>',
    '  padded on both sides  ',
    'line one\nline two',
  ];
  const numbers = [1e21, 0.1 + 0.2, 123456789.123456789, 1.7976931348623157e308];

  describe.each([false, true])('with shared strings %s', sharedStrings => {
    const bytes = new ExcelWriter({ sharedStrings }).createWorkbookBuffer([
      { data: [values, numbers] },
    ]);

    it.each([
      ['written', bytes],
      [
        'after two Workbook round trips',
        Workbook.fromBuffer(Workbook.fromBuffer(bytes).toBuffer()).toBuffer(),
      ],
    ])('reads every string and number back exactly (%s)', (_label, file) => {
      const [first, second] = ExcelBridge.read(file).sheets[0].data;
      expect(first.map(cell => cell.value)).toEqual(values);
      expect(first.every(cell => cell.type === 'string')).toBe(true);
      expect(second.map(cell => cell.value)).toEqual(numbers);
    });
  });
});

describe('what the reader flattens or refuses', () => {
  it('flattens rich text runs into one string', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
      {
        'xl/sharedStrings.xml': strToU8(
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>bold</t></r><r><t xml:space="preserve"> and plain</t></r></si></sst>'
        ),
      }
    );
    expect(cellAt(ExcelBridge.read(bytes).sheets[0], 'A1')?.value).toBe('bold and plain');
  });

  it('refuses a sheet that declares an external entity', () => {
    const doctype =
      '<!DOCTYPE worksheet [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&xxe;</t></is></c></row></sheetData></worksheet>';
    const bytes = buildXlsx('', { 'xl/worksheets/sheet1.xml': strToU8(doctype) });
    expect(() => ExcelBridge.read(bytes)).toThrow(/External entities are not supported/);
  });
});
