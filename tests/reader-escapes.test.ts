import { describe, expect, it } from 'vitest';
import { strToU8 } from 'fflate';
import {
  ExcelReader,
  ExcelWriter,
  Workbook,
  createExcelWorkbookStream,
  type CellValue,
} from '../src';
import { part, readFirstSheet } from './helpers/read';
import { SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const sharedStringsPart = (items: string): Record<string, Uint8Array> => ({
  'xl/sharedStrings.xml': strToU8(`<sst xmlns="${SPREADSHEET_NS}">${items}</sst>`),
});

const valuesOf = (bytes: Uint8Array): unknown[] =>
  new ExcelReader()
    .parseFromBuffer(bytes)
    .sheets[0].data.flat()
    .map(cell => cell.value);

const roundTrip = (text: string, options?: { sharedStrings?: boolean }): unknown => {
  const bytes = new ExcelWriter(options).createWorkbookBuffer([{ data: [[text]] }]);
  return valuesOf(bytes)[0];
};

const streamed = async (text: string): Promise<unknown> => {
  const chunks: Uint8Array[] = [];
  for await (const chunk of createExcelWorkbookStream([{ rows: [[text]] as CellValue[][] }])) {
    chunks.push(chunk);
  }
  const bytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return valuesOf(bytes)[0];
};

describe('character escapes in text read from other tools', () => {
  it.each([
    ['a_x000D_b', 'a\rb'],
    ['a_x000a_b', 'a\nb'],
    ['_x0041__x0042_', 'AB'],
    ['_x00E9_', 'é'],
    ['_xD83D__xDE00_', '😀'],
    ['_x005F_x000D_', '_x000D_'],
    ['_x005F_', '_'],
    ['_x000D__x000A_', '\r\n'],
  ])('decodes %s in an inline string', (escaped, expected) => {
    const body = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${escaped}</t></is></c></row></sheetData>`;
    expect(valuesOf(buildXlsx(body))).toEqual([expected]);
  });

  it.each(['_x00_', '_x000_', '_xZZZZ_', '_x0041', 'x0041_', '__x41__', '_ x0041_'])(
    'leaves %s as it is, because it is not an escape',
    text => {
      const body = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${text}</t></is></c></row></sheetData>`;
      expect(valuesOf(buildXlsx(body))).toEqual([text]);
    }
  );

  it('decodes shared strings, and decodes an escape split across rich text runs once the runs are joined', () => {
    const body =
      '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>';
    const sharedStrings = sharedStringsPart(
      '<si><t>one_x000D_two</t></si><si><r><t>_x00</t></r><r><t>41_</t></r></si>'
    );
    expect(valuesOf(buildXlsx(body, sharedStrings))).toEqual(['one\rtwo', 'A']);
  });

  it('decodes the text result of a formula', () => {
    const body =
      '<sheetData><row r="1"><c r="A1" t="str"><f>"x"</f><v>a_x0009_b</v></c></row></sheetData>';
    const sheet = readFirstSheet(buildXlsx(body));
    expect(sheet.data[0][0]).toMatchObject({ value: 'a\tb', formula: '"x"' });
  });

  it('does not decode the text of a formula', () => {
    const body =
      '<sheetData><row r="1"><c r="A1" t="str"><f>"_x0041_"</f><v>x</v></c></row></sheetData>';
    expect(readFirstSheet(buildXlsx(body)).data[0][0]).toMatchObject({ formula: '"_x0041_"' });
  });

  it('does not decode a sheet name', () => {
    const bytes = buildXlsx('<sheetData/>', {}, { sheetName: 'a_x0041_b' });
    expect(new ExcelReader().parseFromBuffer(bytes).sheets[0].name).toBe('a_x0041_b');
  });
});

describe('text that looks like an escape survives a write and a read', () => {
  const texts = [
    '_x000D_',
    'a_x000D_b',
    '_x005F_',
    '_x005F_x000D_',
    '_x000D__x000A_',
    '__x000D__',
    'x_x0041_y_x0042_',
    '_x00',
    '_X000D_',
    'plain text',
  ];

  it.each(texts)('keeps %j with inline strings', text => {
    expect(roundTrip(text)).toBe(text);
  });

  it.each(texts)('keeps %j with shared strings', text => {
    expect(roundTrip(text, { sharedStrings: true })).toBe(text);
  });

  it.each(texts)('keeps %j in a streamed workbook', async text => {
    expect(await streamed(text)).toBe(text);
  });

  it('keeps the text through a Workbook load and save', () => {
    const bytes = new ExcelWriter().createWorkbookBuffer([{ data: [['_x000D_', 'a_x0041_b']] }]);
    expect(valuesOf(Workbook.fromBuffer(bytes).toBuffer())).toEqual(['_x000D_', 'a_x0041_b']);
  });

  it('writes the underscore of an escape-looking text as _x005F_', () => {
    const bytes = new ExcelWriter().createWorkbookBuffer([{ data: [['_x000D_']] }]);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<t>_x005F_x000D_</t>');
  });

  it('writes any other text exactly as before', () => {
    const bytes = new ExcelWriter().createWorkbookBuffer([{ data: [['a_b x1_ _x1_']] }]);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<t>a_b x1_ _x1_</t>');
  });
});
