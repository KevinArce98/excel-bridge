import { describe, expect, it } from 'vitest';
import { ExcelWriter, Workbook, createExcelWorkbookStream, streamToBuffer } from '../src';
import type { CellStyle, StreamingSheetInput } from '../src';
import { part, readFirstSheet } from './helpers/read';
import { buildXlsx } from './helpers/xlsx';

const writer = new ExcelWriter();
const write = (styles: Record<string, CellStyle>) =>
  writer.createWorkbookBuffer([{ data: [['a', 'b']], styles }]);
const stylesXml = (styles: Record<string, CellStyle>) => part(write(styles), 'xl/styles.xml');
const bordersOf = (xml: string) => /<borders[\s\S]*?<\/borders>/.exec(xml)![0];
const borderCount = (xml: string) => Number(/<borders count="(\d+)"/.exec(xml)![1]);

const THIN_BLACK_BOX =
  '<border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom>\n      <diagonal/>\n    </border>';

describe('border input forms', () => {
  it.each<[string, CellStyle['border']]>([
    ['true', true],
    ['a thin style name', 'thin'],
    ['a thin line', { style: 'thin' }],
    ['an explicit black line', { style: 'thin', color: '#000000' }],
    ['a short black line', { style: 'thin', color: '#000' }],
    ['four thin sides', { left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' }],
  ])('%s writes the box that border: true always wrote', (_, border) => {
    const xml = stylesXml({ '0-0': { border } });
    expect(bordersOf(xml)).toContain(THIN_BLACK_BOX);
    expect(borderCount(xml)).toBe(2);
    expect(xml).toContain('borderId="1" xfId="0" applyBorder="1"');
  });

  it.each<[CellStyle['border']]>([[false], [undefined], [{}]])('%j writes no border', border => {
    const xml = stylesXml({ '0-0': { border } });
    expect(borderCount(xml)).toBe(1);
    expect(xml).not.toContain('applyBorder');
  });

  it('writes sides in schema order whatever the key order', () => {
    const xml = stylesXml({ '0-0': { border: { bottom: 'double', left: 'hair' } } });
    expect(bordersOf(xml)).toContain(
      '<border><left style="hair"><color rgb="FF000000"/></left><right/><top/><bottom style="double"><color rgb="FF000000"/></bottom>\n      <diagonal/>'
    );
  });

  it('applies one style to all four sides and keeps the colour', () => {
    const xml = stylesXml({ '0-0': { border: { style: 'medium', color: '#abc' } } });
    expect(bordersOf(xml).match(/style="medium"><color rgb="FFAABBCC"\/>/g)).toHaveLength(4);
  });

  it('colours one side without touching the others', () => {
    const xml = stylesXml({ '0-0': { border: { top: { style: 'thick', color: '#80ff0000' } } } });
    expect(bordersOf(xml)).toContain(
      '<left/><right/><top style="thick"><color rgb="80FF0000"/></top><bottom/>'
    );
  });

  it('shares a border between equal styles written in a different order', () => {
    const xml = stylesXml({
      '0-0': { border: { left: 'thin', bottom: 'thin' } },
      '0-1': { border: { bottom: 'thin', left: { style: 'thin', color: '#000' } } },
      '1-0': { bold: true, border: { left: 'thin', bottom: 'thin' } },
    });
    expect(borderCount(xml)).toBe(2);
  });

  it.each([
    'thin',
    'medium',
    'thick',
    'dashed',
    'dotted',
    'double',
    'hair',
    'mediumDashed',
    'dashDot',
    'mediumDashDot',
    'dashDotDot',
    'mediumDashDotDot',
    'slantDashDot',
  ] as const)('accepts %s and reads it back', style => {
    const sheet = readFirstSheet(write({ '0-0': { border: { right: style } } }));
    expect(sheet.styles?.['0-0']?.border).toEqual({ right: { style } });
  });
});

describe('border validation', () => {
  it.each([
    ['an unknown style', { left: 'wavy' }, /Invalid border style "wavy"/],
    ['none', { left: 'none' }, /Invalid border style "none"/],
    ['a style name for all sides', 'dash', /Invalid border style "dash"/],
    ['a line without a style', { left: {} }, /Invalid border style "\{\}"/],
    ['a boolean side', { left: true }, /Invalid border style "true"/],
    ['a non-boolean truthy value', 1, /Invalid border style/],
    ['a bad colour', { top: { style: 'thin', color: 'red' } }, /Invalid colour "red"/],
    ['a bad colour on every side', { style: 'thin', color: '#12' }, /Invalid colour/],
  ])('rejects %s', (_, border, message) => {
    expect(() => write({ '0-0': { border: border as CellStyle['border'] } })).toThrow(message);
  });

  it('rejects the same input in a streamed sheet before pulling the first row', async () => {
    let pulled = 0;
    function* rows() {
      pulled++;
      yield ['a'];
    }
    const sheet: StreamingSheetInput = {
      rows: rows(),
      styles: { '999-0': { border: { bottom: 'wavy' as never } } },
    };
    await expect(streamToBuffer(createExcelWorkbookStream([sheet]))).rejects.toThrow(
      /Invalid border style "wavy"/
    );
    expect(pulled).toBe(0);
  });
});

describe('streaming writer', () => {
  it('writes the same styles part as ExcelWriter for per-side borders', async () => {
    const styles: Record<string, CellStyle> = {
      '0-0': { border: { bottom: { style: 'double', color: '#FF0000' } } },
      '1-0': { border: 'medium', bold: true },
      '1-1': { border: true },
    };
    const rows = [
      ['a', 'b'],
      ['c', 'd'],
    ];
    const streamed = await streamToBuffer(createExcelWorkbookStream([{ rows, styles }]));
    const written = writer.createWorkbookBuffer([{ data: rows, styles }]);
    expect(part(streamed, 'xl/styles.xml')).toBe(part(written, 'xl/styles.xml'));
  });
});

describe('reading borders', () => {
  const read = (styles: Record<string, CellStyle>) => readFirstSheet(write(styles)).styles;

  it('reads border: true back as true', () => {
    expect(read({ '0-0': { border: true } })?.['0-0']?.border).toBe(true);
  });

  it('reads a thin box in another colour as four lines', () => {
    expect(
      read({ '0-0': { border: { style: 'thin', color: '#FF0000' } } })?.['0-0']?.border
    ).toEqual({
      left: { style: 'thin', color: '#FF0000' },
      right: { style: 'thin', color: '#FF0000' },
      top: { style: 'thin', color: '#FF0000' },
      bottom: { style: 'thin', color: '#FF0000' },
    });
  });

  it('reads a black line without a colour, so a box of them is true', () => {
    expect(
      read({ '0-0': { border: { left: { style: 'medium', color: '#000' } } } })?.['0-0']?.border
    ).toEqual({ left: { style: 'medium' } });
  });

  it('gives each cell its own border object', () => {
    const styles = read({
      '0-0': { border: { top: 'thin' } },
      '0-1': { border: { top: 'thin' } },
    })!;
    expect(styles['0-0'].border).toEqual(styles['0-1'].border);
    expect(styles['0-0'].border).not.toBe(styles['0-1'].border);
  });

  const foreignStyles = (borders: string) =>
    `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="2"><border/><border>${borders}</border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/></cellXfs></styleSheet>`;
  const foreignBorder = (borders: string) =>
    readFirstSheet(
      buildXlsx(
        '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>',
        {},
        { styles: foreignStyles(borders) }
      )
    ).styles?.['0-0']?.border;

  it('treats a style of none and unknown styles as no line', () => {
    expect(foreignBorder('<left style="none"/><right style="wavy"/><top style="thin"/>')).toEqual({
      top: { style: 'thin' },
    });
  });

  it('drops theme, indexed and automatic colours but keeps the line', () => {
    expect(
      foreignBorder(
        '<left style="thin"><color theme="1" tint="-0.25"/></left><right style="hair"><color indexed="10"/></right><top style="thin"><color auto="1"/></top>'
      )
    ).toEqual({ left: { style: 'thin' }, right: { style: 'hair' }, top: { style: 'thin' } });
  });

  it('reads four thin lines in automatic or lower-case black as true', () => {
    expect(
      foreignBorder(
        ['left', 'right', 'top', 'bottom']
          .map(s => `<${s} style="thin"><color indexed="64"/></${s}>`)
          .join('')
      )
    ).toBe(true);
    expect(
      foreignBorder(
        ['left', 'right', 'top', 'bottom']
          .map(s => `<${s} style="thin"><color rgb="ff000000"/></${s}>`)
          .join('')
      )
    ).toBe(true);
  });

  it('ignores a diagonal-only border', () => {
    expect(
      foreignBorder('<diagonal style="thin"><color rgb="FF0000FF"/></diagonal>')
    ).toBeUndefined();
  });
});

describe('Workbook round trip', () => {
  const styles: Record<string, CellStyle> = {
    '0-0': { border: true },
    '0-1': { border: { top: { style: 'medium', color: '#FF0000' }, bottom: 'double' } },
    '1-0': { bold: true, border: { style: 'dashed', color: '#00B050' } },
    '1-1': { border: { left: 'hair', right: 'slantDashDot' } },
  };
  const original = writer.createWorkbookBuffer([
    {
      data: [
        ['a', 'b'],
        ['c', 'd'],
      ],
      styles,
    },
  ]);
  const saved = Workbook.fromBuffer(original).toBuffer();

  it('writes the sheet and the borders part unchanged', () => {
    expect(part(saved, 'xl/worksheets/sheet1.xml')).toBe(
      part(original, 'xl/worksheets/sheet1.xml')
    );
    expect(bordersOf(part(saved, 'xl/styles.xml'))).toBe(
      bordersOf(part(original, 'xl/styles.xml'))
    );
  });

  it('keeps the styles through a second pass', () => {
    const sheet = readFirstSheet(saved);
    expect(sheet.styles?.['0-0']?.border).toBe(true);
    expect(sheet.styles?.['0-1']?.border).toEqual({
      top: { style: 'medium', color: '#FF0000' },
      bottom: { style: 'double' },
    });
  });
});
