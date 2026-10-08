import { describe, expect, it } from 'vitest';
import { ExcelWriter, Workbook, createExcelWorkbookStream, streamToBuffer } from '../src';
import type { CellStyle, CellValue } from '../src';
import { part, readFirstSheet } from './helpers/read';
import { DATE_STYLES, buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();
const day = new Date(2024, 0, 15);

const write = (cell: CellValue, style?: CellStyle) =>
  writer.createWorkbookBuffer([{ data: [[cell]], styles: style ? { '0-0': style } : undefined }]);

describe('styles on Date cells', () => {
  it('keeps the built-in date format for a Date without a style', () => {
    const bytes = write(day);
    expect(part(bytes, 'xl/styles.xml')).toContain(
      '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    );
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="1">');
  });

  it('writes the style of a Date cell next to the built-in date format', () => {
    const styles = part(write(day, { bold: true, background: '#FFFF00' }), 'xl/styles.xml');
    expect(styles).toContain('numFmtId="14"');
    expect(styles).toContain('applyFont="1" applyFill="1" applyNumberFormat="1"');
  });

  it('lets numberFormat replace the built-in date format', () => {
    const bytes = write(day, { numberFormat: 'yyyy-mm-dd hh:mm' });
    expect(part(bytes, 'xl/styles.xml')).toContain('formatCode="yyyy-mm-dd hh:mm"');
    expect(readFirstSheet(bytes).data[0][0].type).toBe('date');
  });

  it('styles the Date result of a formula', () => {
    const styles = part(
      write({ formula: 'TODAY()', result: day }, { bold: true }),
      'xl/styles.xml'
    );
    expect(styles).toContain('numFmtId="14" fontId="1"');
  });

  it('shares one xf between a Date style and the same style on a number', () => {
    const bytes = writer.createWorkbookBuffer([
      {
        data: [[day, 1]],
        styles: {
          '0-0': { numberFormat: 'dd/mm/yyyy' },
          '0-1': { numberFormat: 'dd/mm/yyyy' },
        },
      },
    ]);
    expect(part(bytes, 'xl/styles.xml')).toContain('<cellXfs count="2">');
  });

  it('streams the style of a Date cell', async () => {
    const bytes = await streamToBuffer(
      createExcelWorkbookStream([
        { rows: [[day]], styles: { '0-0': { numberFormat: 'dd/mm/yyyy' } } },
      ])
    );
    expect(part(bytes, 'xl/styles.xml')).toContain('formatCode="dd/mm/yyyy"');
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="1">');
  });
});

describe('Date cells through the reader and Workbook', () => {
  it('reads the style of a Date cell', () => {
    expect(readFirstSheet(write(day, { bold: true })).styles?.['0-0']?.bold).toBe(true);
  });

  it('keeps a custom date format and a font through a save', () => {
    const saved = readFirstSheet(
      Workbook.fromBuffer(write(day, { numberFormat: 'dd/mm/yyyy', bold: true })).toBuffer()
    );
    expect(saved.styles?.['0-0']).toMatchObject({ numberFormat: 'dd/mm/yyyy', bold: true });
    expect(saved.data[0][0].type).toBe('date');
  });

  it('reads no number format for a built-in date format', () => {
    const bytes = buildXlsx(
      '<sheetData><row r="1"><c r="A1" s="1"><v>40000</v></c></row></sheetData>',
      {},
      { styles: DATE_STYLES }
    );
    const sheet = readFirstSheet(bytes);
    expect(sheet.data[0][0].type).toBe('date');
    expect(sheet.styles?.['0-0']?.numberFormat).toBeUndefined();
  });
});
