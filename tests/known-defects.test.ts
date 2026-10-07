import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  ExcelBridge,
  ExcelWriter,
  calculateColumnWidths,
  createExcelWorkbookStream,
  streamToBuffer,
} from '../src';
import { knownDefect } from './helpers/known-defect';
import { calendarDay, cellAt, readFirstSheet } from './helpers/read';
import { DATE_STYLES, REL_NS, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();

const sheetBody = (cells: string) => `<sheetData><row r="1">${cells}</row></sheetData>`;

describe('number formats', () => {
  const formatted = (numberFormat: string) =>
    readFirstSheet(
      writer.createWorkbookBuffer([{ data: [[1]], styles: { '0-0': { numberFormat } } }])
    ).styles?.['0-0']?.numberFormat;

  it.each(['#,##0.00', '0.0%', '$#,##0.00', '0.00', '00000', '0.0', '0.000', '0.00E+00'])(
    'reads %s back unchanged',
    code => {
      expect(formatted(code)).toBe(code);
    }
  );
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

  knownDefect(
    'a Date cell loses its bold style (expected: bold stays) (I4)',
    () => {
      expect(styleOf(new Date(2024, 0, 15))?.bold).toBe(true);
    },
    { message: /expected undefined to be true/ }
  );
});

describe('date systems', () => {
  const serial = sheetBody('<c r="A1" s="1"><v>40000</v></c>');

  it('reads the 1900 system', () => {
    const date = cellAt(readFirstSheet(buildXlsx(serial, {}, { styles: DATE_STYLES })), 'A1');
    expect(calendarDay(date?.value)).toEqual([2009, 7, 6]);
  });

  knownDefect(
    'a date1904 workbook is read with the 1900 epoch (expected 2013-07-07) (R4)',
    () => {
      const bytes = buildXlsx(
        serial,
        {},
        { styles: DATE_STYLES, workbookProperties: '<workbookPr date1904="1"/>' }
      );
      expect(calendarDay(cellAt(readFirstSheet(bytes), 'A1')?.value)).toEqual([2013, 7, 7]);
    },
    { message: /expected \[ 2009, 7, 6 \] to deeply equal \[ 2013, 7, 7 \]/ }
  );
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

  knownDefect(
    'an error cell is read as NaN (expected: a value that is not NaN) (R5)',
    () => {
      expect(Number.isNaN(cellAt(sheet, 'A1')?.value)).toBe(false);
    },
    { message: /expected true to be false/ }
  );

  knownDefect(
    'the escape _x000D_ stays in the text (expected: a carriage return) (R6)',
    () => {
      expect(cellAt(sheet, 'C1')?.value).toBe('a\rb');
    },
    { message: /expected 'a_x000D_b' to be/ }
  );
});

describe('formulas from other tools', () => {
  const body =
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData>';
  const sheet = readFirstSheet(buildXlsx(body));

  it('reads the master of a shared formula', () => {
    expect(cellAt(sheet, 'B1')?.formula).toBe('A1*2');
  });

  knownDefect(
    'a shared formula follower reads an empty formula (expected: A2*2) (R9)',
    () => {
      expect(cellAt(sheet, 'B2')?.formula).toBe('A2*2');
    },
    { message: /expected '' to be 'A2\*2'/ }
  );
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

  knownDefect(
    'a split pane is read as a freeze pane of 1800 rows (expected: no freeze pane) (R13)',
    () => {
      expect(
        paneSheet('<pane xSplit="2400" ySplit="1800" state="split"/>').freezePane
      ).toBeUndefined();
    },
    { message: /expected \{ row: 1800, col: 2400 \} to be undefined/ }
  );
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
    'a workbook with prefixed elements is rejected as invalid (expected: its sheet is read) (R8)',
    () => {
      expect(() => ExcelBridge.read(prefixed('x'))).not.toThrow();
    },
    { message: /not throw an error but .*Failed to parse Excel file/ }
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

  const accepted = { message: /to throw an error/ };

  knownDefect(
    'a sheet name longer than 31 characters is written (expected: an error) (W3)',
    () => {
      expect(() => write([{ data: [['a']], options: { name: 'x'.repeat(32) } }])).toThrow();
    },
    accepted
  );

  knownDefect(
    'a sheet name with a forbidden character is written (expected: an error) (W3)',
    () => {
      expect(() => write([{ data: [['a']], options: { name: 'bad:name' } }])).toThrow();
    },
    accepted
  );

  knownDefect(
    'two sheets with the same name are written (expected: an error) (W3)',
    () => {
      expect(() =>
        write([
          { data: [['a']], options: { name: 'Same' } },
          { data: [['b']], options: { name: 'Same' } },
        ])
      ).toThrow();
    },
    accepted
  );

  knownDefect(
    'a workbook without sheets is written (expected: an error) (W3)',
    () => {
      expect(() => write([])).toThrow();
    },
    accepted
  );

  knownDefect(
    'a colour that is not hex is written as an invalid ARGB value (expected: an error) (W4)',
    () => {
      expect(() => write([{ data: [['a']], styles: { '0-0': { color: 'red' } } }])).toThrow();
    },
    accepted
  );

  knownDefect(
    'NaN is written as a number cell (expected: an error or an error cell, decided with the writer validation) (R5)',
    () => {
      expect(() => write([{ data: [[NaN]] }])).toThrow();
    },
    accepted
  );

  knownDefect(
    'Infinity is written as a number cell (expected: an error or an error cell, decided with the writer validation) (R5)',
    () => {
      expect(() => write([{ data: [[Infinity]] }])).toThrow();
    },
    accepted
  );
});

describe('column width measurement on large sheets', () => {
  const tall = (rows: number) => Array.from({ length: rows }, () => ['x']);

  it('measures a thousand rows', () => {
    expect(calculateColumnWidths(tall(1000))).toHaveLength(1);
  });

  knownDefect(
    'measuring one million rows overflows the call stack (expected: a width) (W2)',
    () => calculateColumnWidths(tall(1_000_000)),
    { error: 'RangeError', message: /Maximum call stack size exceeded/ }
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

  knownDefect(
    'most streamed entries are stored without compression, which SheetJS 0.18.5 cannot read (expected: all deflated) (W1)',
    async () => {
      const methods = await compressionMethods();
      expect(methods).toEqual(methods.map(() => 8));
    },
    { message: /to deeply equal/ }
  );
});
