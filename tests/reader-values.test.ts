import { describe, expect, it } from 'vitest';
import { ExcelBridge, Workbook } from '../src';
import { cellAt, readFirstSheet } from './helpers/read';
import { DATE_STYLES, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const row = (cells: string) => `<sheetData><row r="1">${cells}</row></sheetData>`;

const roundTrip = (bytes: Uint8Array) => Workbook.fromBuffer(bytes).toBuffer();

describe('error cells and values a worksheet cannot hold', () => {
  const bytes = buildXlsx(
    row(
      [
        '<c r="A1" t="e"><v>#DIV/0!</v></c>',
        '<c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
        '<c r="C1"><v>1e999</v></c>',
        '<c r="D1"><v>NaN</v></c>',
        '<c r="E1" s="1"><v>1e12</v></c>',
        '<c r="F1"><v>7</v></c>',
      ].join('')
    ),
    {},
    { styles: DATE_STYLES }
  );
  const sheet = readFirstSheet(bytes);

  it('reads an error cell as its error text', () => {
    expect(cellAt(sheet, 'A1')).toMatchObject({ type: 'error', value: '#DIV/0!' });
  });

  it('keeps the formula of a formula cell that holds an error', () => {
    expect(cellAt(sheet, 'B1')).toMatchObject({ type: 'error', value: '#DIV/0!', formula: '1/0' });
  });

  it.each(['C1', 'D1'])('reads the non-finite number in %s as #NUM!', coordinate => {
    expect(cellAt(sheet, coordinate)).toMatchObject({ type: 'error', value: '#NUM!' });
  });

  it('reads a date-formatted serial that is not a date as a number', () => {
    expect(cellAt(sheet, 'E1')).toMatchObject({ type: 'number', value: 1e12 });
  });

  it('leaves ordinary numbers alone', () => {
    expect(cellAt(sheet, 'F1')).toMatchObject({ type: 'number', value: 7 });
  });

  it('saves a workbook that holds them instead of throwing', () => {
    const saved = readFirstSheet(roundTrip(bytes));
    expect(cellAt(saved, 'A1')).toMatchObject({ type: 'string', value: '#DIV/0!' });
    expect(cellAt(saved, 'B1')?.formula).toBe('1/0');
    expect(cellAt(saved, 'C1')).toMatchObject({ type: 'string', value: '#NUM!' });
    expect(cellAt(saved, 'D1')).toMatchObject({ type: 'string', value: '#NUM!' });
    expect(cellAt(saved, 'E1')).toMatchObject({ type: 'number', value: 1e12 });
    expect(cellAt(saved, 'F1')?.value).toBe(7);
  });
});

describe('wrap text from other tools', () => {
  const styles = (wrapText: string) =>
    `<?xml version="1.0"?><styleSheet xmlns="${SPREADSHEET_NS}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment wrapText="${wrapText}"/></xf></cellXfs></styleSheet>`;

  it.each([
    ['1', true],
    ['true', true],
    ['0', undefined],
    ['false', undefined],
  ])('reads wrapText="%s"', (wrapText, expected) => {
    const bytes = buildXlsx(row('<c r="A1" s="1"><v>1</v></c>'), {}, { styles: styles(wrapText) });
    expect(ExcelBridge.read(bytes).sheets[0].styles?.['0-0']?.wrapText).toBe(expected);
  });
});
