import { describe, expect, it } from 'vitest';
import { ExcelBridge, ExcelWriter, Workbook, dataValidation, hyperlink } from '../src';
import type { CellValidation } from '../src';
import { cellAt, part, readFirstSheet } from './helpers/read';
import { buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();

const roundTrip = (bytes: Uint8Array): Uint8Array => Workbook.fromBuffer(bytes).toBuffer();

describe('rows keep their position through Workbook', () => {
  const gapped = (() => {
    const workbook = Workbook.create();
    workbook.addSheet('S', [['a']]);
    workbook.setCellValue('S', 4, 0, 'e');
    workbook.setCellStyle('S', 4, 0, { bold: true });
    return workbook.toBuffer();
  })();

  it.each([
    ['the first save', gapped],
    ['a second save', roundTrip(gapped)],
    ['a third save', roundTrip(roundTrip(gapped))],
  ])('puts a row gap back where it was after %s', (_label, bytes) => {
    const sheet = readFirstSheet(bytes);
    expect(cellAt(sheet, 'A1')?.value).toBe('a');
    expect(cellAt(sheet, 'A5')?.value).toBe('e');
    expect(sheet.styles?.['4-0']?.bold).toBe(true);
  });

  it('reads and writes a cell at the position getCellValue uses', () => {
    const workbook = Workbook.fromBuffer(gapped);
    expect(workbook.getCellValue('S', 4, 0)).toBe('e');
    expect(workbook.getCellValue('S', 2, 0)).toBeNull();
  });

  it('keeps an empty row element from shifting the rows after it', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"/><row r="3"><c r="A3"><v>3</v></c></row></sheetData>'
    );
    expect(cellAt(readFirstSheet(roundTrip(bytes)), 'A3')?.value).toBe(3);
  });

  it('keeps rows that arrive out of order instead of overwriting one', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="3"><c r="A3"><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>'
    );
    const values = readFirstSheet(roundTrip(bytes))
      .data.flat()
      .map(cell => cell.value);
    expect(values.sort()).toEqual([2, 3]);
  });

  it('numbers rows that lack an r attribute after the previous row', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="2"><c r="A2"><v>2</v></c></row><row><c r="A3"><v>3</v></c></row></sheetData>'
    );
    const sheet = readFirstSheet(roundTrip(bytes));
    expect(cellAt(sheet, 'A2')?.value).toBe(2);
    expect(cellAt(sheet, 'A3')?.value).toBe(3);
  });
});

describe('data validations keep their rule through Workbook', () => {
  const rules: Array<[string, CellValidation, RegExp]> = [
    [
      'whole number',
      dataValidation.wholeNumber('A2:A10', 'between', 1, 100),
      /type="whole" operator="between"[^>]*sqref="A2:A10">\s*<formula1>1<\/formula1><formula2>100<\/formula2>/,
    ],
    [
      'decimal',
      dataValidation.decimal('B2:B10', 'greaterThanOrEqual', 0.5),
      /type="decimal" operator="greaterThanOrEqual"[^>]*sqref="B2:B10">\s*<formula1>0\.5<\/formula1>/,
    ],
    [
      'text length',
      dataValidation.textLength('C2:C10', 'lessThanOrEqual', 20),
      /type="textLength" operator="lessThanOrEqual"[^>]*sqref="C2:C10">\s*<formula1>20<\/formula1>/,
    ],
    [
      'date',
      dataValidation.dateBetween('D2:D10', new Date(2024, 0, 1), new Date(2024, 11, 31)),
      /type="date" operator="between"[^>]*sqref="D2:D10">\s*<formula1>45292<\/formula1><formula2>45657<\/formula2>/,
    ],
    [
      'time',
      { range: 'E2:E10', type: 'time', operator: 'lessThan', formula1: '0.5', options: '' },
      /type="time" operator="lessThan"[^>]*sqref="E2:E10">\s*<formula1>0\.5<\/formula1>/,
    ],
    [
      'custom formula',
      { range: 'F2:F10', type: 'custom', formula1: 'ISNUMBER(F2)', options: '' },
      /type="custom" allowBlank="1"[^>]*sqref="F2:F10">\s*<formula1>ISNUMBER\(F2\)<\/formula1>/,
    ],
    [
      'inline list',
      dataValidation.list('G2:G10', ['Open', 'Done']),
      /type="list"[^>]*sqref="G2:G10">\s*<formula1>"Open,Done"<\/formula1>/,
    ],
    [
      'list from a range',
      { range: 'H2:H10', type: 'list', formula1: '$K$1:$K$3', options: '' },
      /type="list"[^>]*sqref="H2:H10">\s*<formula1>\$K\$1:\$K\$3<\/formula1>/,
    ],
    [
      'inline list with quotes',
      dataValidation.list('I2:I10', ['say "hi"', 'bye']),
      /type="list"[^>]*sqref="I2:I10">\s*<formula1>"say ""hi"",bye"<\/formula1>/,
    ],
  ];

  const original = writer.createWorkbookBuffer([
    { data: [['x']], validations: rules.map(([, validation]) => validation) },
  ]);
  const saved = roundTrip(original);

  it.each(rules)('writes the %s rule', (_label, _validation, pattern) => {
    expect(part(original, 'xl/worksheets/sheet1.xml')).toMatch(pattern);
  });

  it.each(rules)('keeps the %s rule after a save', (_label, _validation, pattern) => {
    expect(part(saved, 'xl/worksheets/sheet1.xml')).toMatch(pattern);
  });

  it('writes the whole validations block unchanged', () => {
    const block = (bytes: Uint8Array) =>
      /<dataValidations[\s\S]*<\/dataValidations>/.exec(
        part(bytes, 'xl/worksheets/sheet1.xml')
      )?.[0];
    expect(block(saved)).toBe(block(original));
  });

  it('keeps allowBlank false', () => {
    const bytes = writer.createWorkbookBuffer([
      {
        data: [['x']],
        validations: [{ ...dataValidation.wholeNumber('A2', 'equal', 1), allowBlank: false }],
      },
    ]);
    expect(part(roundTrip(bytes), 'xl/worksheets/sheet1.xml')).toMatch(/allowBlank="0"/);
  });

  it('reads the rule back with its type, operator and formulas', () => {
    expect(readFirstSheet(original).validations[0]).toMatchObject({
      range: 'A2:A10',
      type: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '100',
      allowBlank: true,
    });
  });

  it('reads an inline list as its values without the quotes', () => {
    const inline = readFirstSheet(original).validations.find(rule => rule.range === 'I2:I10');
    expect(inline).toMatchObject({ type: 'list', options: 'say "hi",bye' });
    expect(inline?.formula1).toBeUndefined();
  });

  it('drops a rule of an unsupported type instead of rewriting it', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><dataValidations count="1"><dataValidation type="none" sqref="A1" promptTitle="Hint"/></dataValidations>'
    );
    expect(readFirstSheet(bytes).validations).toEqual([]);
  });
});

describe('text that looks like a number stays text', () => {
  it('keeps a sheet name, a title and link texts that start with zeros', () => {
    const original = new ExcelWriter({ title: '007' }).createWorkbookBuffer([
      {
        data: [['link']],
        hyperlinks: [
          hyperlink.url('A1', 'https://example.com/', { tooltip: '007', display: '1.50' }),
        ],
        options: { name: '007' },
      },
    ]);
    const workbook = ExcelBridge.read(original);
    expect(workbook.sheets[0].name).toBe('007');
    expect(workbook.metadata.title).toBe('007');
    expect(workbook.sheets[0].hyperlinks?.[0]).toMatchObject({ tooltip: '007', display: '1.50' });
    expect(Workbook.fromBuffer(original).getSheetNames()).toEqual(['007']);
  });

  it('keeps a validation formula and a rule formula written with trailing zeros', () => {
    const bytes = writer.createWorkbookBuffer([
      {
        data: [['x']],
        validations: [{ range: 'A1', type: 'custom', formula1: '1.50', options: '' }],
        conditionalFormats: [
          { type: 'expression', range: 'A1', formula: '007', style: { bold: true } },
        ],
      },
    ]);
    const sheet = readFirstSheet(bytes);
    expect(sheet.validations[0].formula1).toBe('1.50');
    expect(sheet.conditionalFormats?.[0]).toMatchObject({ formula: '007' });
  });

  it('keeps a colour whose digits are all numeric', () => {
    const styles = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><color rgb="00001234"/></font></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>`;
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>',
      {},
      { styles }
    );
    expect(readFirstSheet(bytes).styles?.['0-0']?.color).toBe('#001234');
  });
});

describe('sheet visibility', () => {
  const build = (states: Array<'visible' | 'hidden' | 'veryHidden' | undefined>) =>
    writer.createWorkbookBuffer(
      states.map((state, index) => ({
        data: [[`s${index}`]],
        options: { name: `S${index}`, state },
      }))
    );

  it('writes the state of each sheet', () => {
    const xml = part(build([undefined, 'hidden', 'veryHidden', 'visible']), 'xl/workbook.xml');
    expect(xml).toMatch(/name="S0" sheetId="1" r:id="rId1"/);
    expect(xml).toMatch(/name="S1" sheetId="2" state="hidden"/);
    expect(xml).toMatch(/name="S2" sheetId="3" state="veryHidden"/);
    expect(xml).toMatch(/name="S3" sheetId="4" r:id="rId4"/);
    expect(xml).not.toMatch(/activeTab/);
  });

  it('reads the state of hidden sheets and leaves visible sheets without one', () => {
    const sheets = ExcelBridge.read(build([undefined, 'hidden', 'veryHidden'])).sheets;
    expect(sheets.map(sheet => sheet.state)).toEqual([undefined, 'hidden', 'veryHidden']);
  });

  it('keeps hidden and very hidden sheets through a save', () => {
    const saved = roundTrip(build([undefined, 'hidden', 'veryHidden']));
    const xml = part(saved, 'xl/workbook.xml');
    expect(xml).toMatch(/name="S1"[^>]*state="hidden"/);
    expect(xml).toMatch(/name="S2"[^>]*state="veryHidden"/);
  });

  it('opens on the first visible sheet when the first sheet is hidden', () => {
    expect(part(build(['hidden', 'visible']), 'xl/workbook.xml')).toMatch(
      /<workbookView[^>]*activeTab="1"/
    );
  });

  it('refuses a workbook where every sheet is hidden', () => {
    expect(() => build(['hidden', 'veryHidden'])).toThrow('At least one sheet must be visible');
  });

  it('changes the state through Workbook', () => {
    const workbook = Workbook.create();
    workbook.addSheet('A', [['a']]);
    workbook.addSheet('B', [['b']]);
    expect(workbook.getSheetState('B')).toBe('visible');

    workbook.setSheetState('B', 'veryHidden');
    expect(workbook.getSheetState('B')).toBe('veryHidden');
    expect(part(workbook.toBuffer(), 'xl/workbook.xml')).toMatch(/name="B"[^>]*state="veryHidden"/);

    workbook.setSheetState('B', 'visible');
    expect(part(workbook.toBuffer(), 'xl/workbook.xml')).not.toMatch(/state=/);
  });
});

describe('number formats keep their code through Workbook', () => {
  it.each(['0.00', '00000', '0.0%', '#,##0.00', '0.00E+00'])('keeps %s', numberFormat => {
    const bytes = writer.createWorkbookBuffer([
      { data: [[1]], styles: { '0-0': { numberFormat } } },
    ]);
    expect(readFirstSheet(roundTrip(bytes)).styles?.['0-0']?.numberFormat).toBe(numberFormat);
  });
});
