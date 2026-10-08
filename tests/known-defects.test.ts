import { describe, expect, it } from 'vitest';
import { ExcelWriter, Workbook } from '../src';
import { knownDefect } from './helpers/known-defect';
import { calendarDay, cellAt, dateOf, readFirstSheet } from './helpers/read';
import { DATE_STYLES, buildXlsx } from './helpers/xlsx';

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

  it('reads the style of a Date cell', () => {
    expect(styleOf(new Date(2024, 0, 15))?.bold).toBe(true);
  });
});

describe('date systems', () => {
  const serial = sheetBody('<c r="A1" s="1"><v>40000</v></c>');

  it('reads the 1900 system', () => {
    const date = cellAt(readFirstSheet(buildXlsx(serial, {}, { styles: DATE_STYLES })), 'A1');
    expect(calendarDay(dateOf(date))).toEqual([2009, 7, 6]);
  });

  it.each(['1', 'true'])('reads the 1904 system when date1904 is %s', flag => {
    const bytes = buildXlsx(
      serial,
      {},
      { styles: DATE_STYLES, workbookProperties: `<workbookPr date1904="${flag}"/>` }
    );
    expect(calendarDay(dateOf(cellAt(readFirstSheet(bytes), 'A1')))).toEqual([2013, 7, 7]);
  });

  it.each(['0', 'false'])('reads the 1900 system when date1904 is %s', flag => {
    const bytes = buildXlsx(
      serial,
      {},
      { styles: DATE_STYLES, workbookProperties: `<workbookPr date1904="${flag}"/>` }
    );
    expect(calendarDay(dateOf(cellAt(readFirstSheet(bytes), 'A1')))).toEqual([2009, 7, 6]);
  });

  it('leaves a plain number alone in the 1904 system and writes the dates back in the 1900 system', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1"><v>40000</v></c></row></sheetData>',
      {},
      { styles: DATE_STYLES, workbookProperties: '<workbookPr date1904="1"/>' }
    );
    const saved = readFirstSheet(Workbook.fromBuffer(bytes).toBuffer());
    expect(calendarDay(dateOf(cellAt(saved, 'A1')))).toEqual([1904, 1, 1]);
    expect(cellAt(saved, 'B1')?.value).toBe(40000);
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

  it('decodes the escape _x000D_ to a carriage return', () => {
    expect(cellAt(sheet, 'C1')?.value).toBe('a\rb');
  });
});

describe('formulas from other tools', () => {
  const body =
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData>';
  const sheet = readFirstSheet(buildXlsx(body));

  it('reads the master of a shared formula', () => {
    expect(cellAt(sheet, 'B1')?.formula).toBe('A1*2');
  });

  it('reads a shared formula follower as its cached value, without a formula', () => {
    expect(cellAt(sheet, 'B2')).toMatchObject({ type: 'number', value: 4 });
    expect(cellAt(sheet, 'B2')).not.toHaveProperty('formula');
  });

  it('saves a workbook with a shared formula follower as its cached value', () => {
    const saved = readFirstSheet(Workbook.fromBuffer(buildXlsx(body)).toBuffer());
    expect(cellAt(saved, 'B2')).toMatchObject({ type: 'number', value: 4 });
    expect(cellAt(saved, 'B2')).not.toHaveProperty('formula');
    expect(cellAt(saved, 'B1')).toMatchObject({ type: 'number', value: 2, formula: 'A1*2' });
  });

  knownDefect(
    'a shared formula follower reads no formula (expected: A2*2) (R9)',
    () => {
      expect(cellAt(sheet, 'B2')?.formula).toBe('A2*2');
    },
    { message: /expected undefined to be 'A2\*2'/ }
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

  it('reads a frozenSplit pane as a freeze pane', () => {
    expect(paneSheet('<pane ySplit="2" state="frozenSplit"/>').freezePane).toEqual({ row: 2 });
  });

  it.each([
    '<pane xSplit="2400" ySplit="1800" state="split"/>',
    '<pane xSplit="2400" ySplit="1800"/>',
  ])('does not read %s as a freeze pane, because its numbers are twips', pane => {
    expect(paneSheet(pane).freezePane).toBeUndefined();
  });
});
