import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  ExcelBridge,
  ExcelWriter,
  Workbook,
  createExcelWorkbookStream,
  dataValidation,
  streamToBuffer,
} from '../src';
import type { ParsedSheet } from '../src/reader';
import { knownDefect } from './helpers/known-defect';
import { DATE_STYLES, REL_NS, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();

const cellAt = (sheet: ParsedSheet, coordinate: string) =>
  sheet.data.flat().find(cell => cell?.coordinate === coordinate);

const calendarDay = (date: Date) => [date.getFullYear(), date.getMonth() + 1, date.getDate()];

const part = (bytes: Uint8Array, path: string): string => strFromU8(unzipSync(bytes)[path]);

const readFirstSheet = (bytes: Uint8Array): ParsedSheet => ExcelBridge.read(bytes).sheets[0];

const sheetBody = (cells: string) => `<sheetData><row r="1">${cells}</row></sheetData>`;

describe('Workbook round trip of the library own output', () => {
  const validated = writer.createWorkbookBuffer([
    { data: [['n']], validations: [dataValidation.wholeNumber('A2:A10', 'between', 1, 100)] },
  ]);

  it('writes a whole-number validation', () => {
    expect(part(validated, 'xl/worksheets/sheet1.xml')).toMatch(/<dataValidation[^>]*type="whole"/);
  });

  knownDefect('I2: a whole-number validation keeps its type', () => {
    const saved = Workbook.fromBuffer(validated).toBuffer();
    expect(part(saved, 'xl/worksheets/sheet1.xml')).toMatch(/<dataValidation[^>]*type="whole"/);
  });

  const gapped = (() => {
    const workbook = Workbook.create();
    workbook.addSheet('S', [['a']]);
    workbook.setCellValue('S', 4, 0, 'e');
    return workbook.toBuffer();
  })();

  it('writes a cell past the last row at its own row', () => {
    expect(cellAt(readFirstSheet(gapped), 'A5')?.value).toBe('e');
  });

  knownDefect('I1: a row gap created by setCellValue survives a second round trip', () => {
    const saved = Workbook.fromBuffer(gapped).toBuffer();
    expect(cellAt(readFirstSheet(saved), 'A5')?.value).toBe('e');
  });
});

describe('number formats', () => {
  const formatted = (numberFormat: string) =>
    readFirstSheet(
      writer.createWorkbookBuffer([{ data: [[1]], styles: { '0-0': { numberFormat } } }])
    ).styles?.['0-0']?.numberFormat;

  it.each(['#,##0.00', '0.0%', '$#,##0.00'])('reads %s back unchanged', code => {
    expect(formatted(code)).toBe(code);
  });

  knownDefect('I3: the format 0.00 is read back unchanged', () => {
    expect(formatted('0.00')).toBe('0.00');
  });

  knownDefect('I3: the format 00000 is read back unchanged', () => {
    expect(formatted('00000')).toBe('00000');
  });
});

describe('styled cells', () => {
  const styleOf = (value: unknown) =>
    readFirstSheet(
      writer.createWorkbookBuffer([
        { data: [[value as number]], styles: { '0-0': { bold: true } } },
      ])
    ).styles?.['0-0'];

  it('reads the style of a number cell', () => {
    expect(styleOf(1)?.bold).toBe(true);
  });

  knownDefect('I4: a Date cell keeps its bold style', () => {
    expect(styleOf(new Date(2024, 0, 15))?.bold).toBe(true);
  });
});

describe('date systems', () => {
  const serial = sheetBody('<c r="A1" s="1"><v>40000</v></c>');

  it('reads the 1900 system', () => {
    const date = cellAt(readFirstSheet(buildXlsx(serial, {}, { styles: DATE_STYLES })), 'A1');
    expect(calendarDay(date?.value)).toEqual([2009, 7, 6]);
  });

  knownDefect('R4: the 1904 system is honoured', () => {
    const bytes = buildXlsx(
      serial,
      {},
      { styles: DATE_STYLES, workbookProperties: '<workbookPr date1904="1"/>' }
    );
    expect(calendarDay(cellAt(readFirstSheet(bytes), 'A1')?.value)).toEqual([2013, 7, 7]);
  });
});

describe('cell values from other tools', () => {
  const body = sheetBody(
    '<c r="A1" t="e"><v>#DIV/0!</v></c><c r="B1"><v>7</v></c><c r="C1" t="inlineStr"><is><t>a_x000D_b</t></is></c><c r="D1" t="inlineStr"><is><t>ok</t></is></c>'
  );
  const sheet = readFirstSheet(buildXlsx(body));

  it('reads the neighbouring cells', () => {
    expect(cellAt(sheet, 'B1')?.value).toBe(7);
    expect(cellAt(sheet, 'D1')?.value).toBe('ok');
  });

  knownDefect('R5: an error cell is not read as NaN', () => {
    expect(Number.isNaN(cellAt(sheet, 'A1')?.value)).toBe(false);
  });

  knownDefect('R6: _x000D_ is decoded to a carriage return', () => {
    expect(cellAt(sheet, 'C1')?.value).toBe('a\rb');
  });
});

describe('formulas from other tools', () => {
  const body = sheetBody(
    '<c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>'
  ).replace(
    '</row>',
    '</row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>'
  );
  const sheet = readFirstSheet(buildXlsx(body));

  it('reads the master of a shared formula', () => {
    expect(cellAt(sheet, 'B1')?.formula).toBe('A1*2');
  });

  knownDefect('R9: a shared formula follower keeps a formula', () => {
    expect(cellAt(sheet, 'B2')?.formula).not.toBe('');
  });
});

describe('sheet views', () => {
  const paneSheet = (pane: string) =>
    readFirstSheet(
      buildXlsx(
        `<sheetViews><sheetView>${pane}</sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`
      )
    );

  it('reads a frozen pane', () => {
    expect(paneSheet('<pane xSplit="1" ySplit="1" state="frozen"/>').freezePane).toEqual({
      row: 1,
      col: 1,
    });
  });

  knownDefect('R13: a split pane is not a freeze pane', () => {
    expect(
      paneSheet('<pane xSplit="2400" ySplit="1800" state="split"/>').freezePane
    ).toBeUndefined();
  });
});

describe('prefixed SpreadsheetML', () => {
  const prefixed = (prefix: string) => {
    const declaration = prefix
      ? `xmlns:${prefix}="${SPREADSHEET_NS}"`
      : `xmlns="${SPREADSHEET_NS}"`;
    const tag = (name: string) => (prefix ? `${prefix}:${name}` : name);
    return zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'xl/workbook.xml': strToU8(
        `<${tag('workbook')} ${declaration} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><${tag('sheets')}><${tag('sheet')} name="S" sheetId="1" r:id="rId1"/></${tag('sheets')}></${tag('workbook')}>`
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        `<${tag('worksheet')} ${declaration}><${tag('sheetData')}><${tag('row')} r="1"><${tag('c')} r="A1"><${tag('v')}>5</${tag('v')}></${tag('c')}></${tag('row')}></${tag('sheetData')}></${tag('worksheet')}>`
      ),
    });
  };

  it('reads the unprefixed equivalent', () => {
    expect(cellAt(readFirstSheet(prefixed('')), 'A1')?.value).toBe(5);
  });

  knownDefect(
    'R8: a prefixed workbook yields its sheet',
    () => {
      const { sheets } = ExcelBridge.read(prefixed('x'));
      expect(sheets).toHaveLength(1);
      expect(cellAt(sheets[0], 'A1')?.value).toBe(5);
    },
    'Error'
  );
});

describe('writer input validation', () => {
  const write = (sheets: Parameters<ExcelWriter['createWorkbookBuffer']>[0]) =>
    writer.createWorkbookBuffer(sheets);

  it('writes a valid sheet name and a hex colour', () => {
    expect(() =>
      write([
        { data: [['a']], options: { name: 'Totals' }, styles: { '0-0': { color: '#FF0000' } } },
      ])
    ).not.toThrow();
  });

  knownDefect('W3: a sheet name longer than 31 characters is rejected', () => {
    expect(() => write([{ data: [['a']], options: { name: 'x'.repeat(32) } }])).toThrow();
  });

  knownDefect('W3: a sheet name with a forbidden character is rejected', () => {
    expect(() => write([{ data: [['a']], options: { name: 'bad:name' } }])).toThrow();
  });

  knownDefect('W3: two sheets with the same name are rejected', () => {
    expect(() =>
      write([
        { data: [['a']], options: { name: 'Same' } },
        { data: [['b']], options: { name: 'Same' } },
      ])
    ).toThrow();
  });

  knownDefect('W3: a workbook without sheets is rejected', () => {
    expect(() => write([])).toThrow();
  });

  knownDefect('W4: a colour that is not hex is rejected', () => {
    expect(() => write([{ data: [['a']], styles: { '0-0': { color: 'red' } } }])).toThrow();
  });

  knownDefect('R5: NaN is rejected instead of written as a number', () => {
    expect(() => write([{ data: [[NaN]] }])).toThrow();
  });

  knownDefect('R5: Infinity is rejected instead of written as a number', () => {
    expect(() => write([{ data: [[Infinity]] }])).toThrow();
  });
});

describe('autoWidth on large sheets', () => {
  const tall = (rows: number) => ({
    data: Array.from({ length: rows }, () => ['x']),
    options: { autoWidth: true },
  });

  it('measures a thousand rows', () => {
    expect(() => writer.createWorkbookBuffer([tall(1000)])).not.toThrow();
  });

  knownDefect(
    'W2: autoWidth handles 150,000 rows',
    () => writer.createWorkbookBuffer([tall(150_000)]),
    'RangeError'
  );
});

describe('streaming writer zip entries', () => {
  const streamed = streamToBuffer(createExcelWorkbookStream([{ name: 'S', rows: [['a', 1]] }]));

  const compressionMethods = async () => {
    const methods: number[] = [];
    unzipSync(await streamed, {
      filter: file => {
        methods.push(file.compression);
        return false;
      },
    });
    return methods;
  };

  it('is readable by the library reader', async () => {
    expect(cellAt(readFirstSheet(await streamed), 'B1')?.value).toBe(1);
  });

  it('lists its entries', async () => {
    expect((await compressionMethods()).length).toBeGreaterThan(0);
  });

  knownDefect('W1: every entry is deflated', async () => {
    expect((await compressionMethods()).every(method => method === 8)).toBe(true);
  });
});
