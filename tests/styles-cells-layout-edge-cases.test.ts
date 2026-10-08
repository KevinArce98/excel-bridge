import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8, zipSync, strToU8 } from 'fflate';
import {
  ExcelWriter,
  ExcelReader,
  Workbook,
  createExcelWorkbookStream,
  generateColsXml,
} from '../src';
import type { CellStyle } from '../src';

const writer = new ExcelWriter();
const part = (buffer: Uint8Array, name: string): string => strFromU8(unzipSync(buffer)[name]);
const sheetXml = (styles: Record<string, any>, data: any[][] = [['a']]) =>
  part(writer.createWorkbookBuffer([{ data, styles }]), 'xl/worksheets/sheet1.xml');
const stylesXml = (buffer: Uint8Array) => part(buffer, 'xl/styles.xml');

describe('style identity', () => {
  it('keeps two class instances with different getters apart', () => {
    class Header {
      get bold() {
        return true;
      }
      get background() {
        return '#FF0000';
      }
    }
    class Total {
      get italic() {
        return true;
      }
      get color() {
        return '#00FF00';
      }
    }
    const buffer = writer.createWorkbookBuffer([
      {
        data: [['a', 'b']],
        styles: { '0-0': new Header() as CellStyle, '0-1': new Total() as CellStyle },
      },
    ]);
    const xml = part(buffer, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" s="1"');
    expect(xml).toContain('<c r="B1" s="2"');
  });

  it('keeps prototype-inherited styles apart and ignores unknown properties', () => {
    const a = Object.create({ bold: true });
    const b = Object.create({ italic: true });
    const circular: any = { bold: true };
    circular.self = circular;
    const xml = sheetXml({ '0-0': a, '0-1': b, '0-2': circular }, [['a', 'b', 'c']]);
    expect(xml).toContain('<c r="A1" s="1"');
    expect(xml).toContain('<c r="B1" s="2"');
    expect(xml).toContain('<c r="C1" s="1"');
    expect(() => sheetXml({ '0-0': { bold: true, extra: 10n } as any })).not.toThrow();
  });
});

describe('falsy styles', () => {
  it.each([false, 0, '', NaN])('writes no s attribute for %s', falsy => {
    expect(sheetXml({ '0-0': falsy })).toContain('<c r="A1" t="inlineStr">');
  });

  it('does the same in the stream', async () => {
    const chunks: Uint8Array[] = [];
    for await (const chunk of createExcelWorkbookStream([
      { rows: [['a']], styles: { '0-0': false as any } },
    ]))
      chunks.push(chunk);
    const total = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    chunks.forEach(c => {
      total.set(c, at);
      at += c.length;
    });
    expect(part(total, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" t="inlineStr">');
  });
});

describe('style numberFormat over a Date', () => {
  const date = new Date(2024, 0, 15, 18, 0);
  const xfs = (numberFormat: string) =>
    stylesXml(
      writer.createWorkbookBuffer([{ data: [[date]], styles: { '0-0': { numberFormat } } }])
    );

  it.each(['@', '#,##0', '0.00'])('keeps the built-in date format for %s', numberFormat => {
    expect(xfs(numberFormat)).toContain('<xf numFmtId="14"');
  });

  it('replaces it by a date format code', () => {
    const xml = xfs('yyyy-mm-dd hh:mm');
    expect(xml).toContain('formatCode="yyyy-mm-dd hh:mm"');
    expect(xml).not.toContain('<xf numFmtId="14"');
  });
});

const dateFile = (): Uint8Array =>
  writer.createWorkbookBuffer([{ data: [['h', new Date(2024, 0, 2), 5]] }]);

describe('reader and Workbook with date cells', () => {
  it('reports no style for a plain date written by 1.5.0 and keeps the round trip bytes', () => {
    const original = dateFile();
    expect(new ExcelReader().parseFromBuffer(original).sheets[0].styles).toBeUndefined();
    const saved = Workbook.fromBuffer(original).toBuffer();
    expect(stylesXml(saved)).toBe(stylesXml(original));
    expect(part(saved, 'xl/worksheets/sheet1.xml')).toBe(
      part(original, 'xl/worksheets/sheet1.xml')
    );
  });

  it('keeps bold on a date', () => {
    const buffer = writer.createWorkbookBuffer([
      { data: [[new Date(2024, 0, 2)]], styles: { '0-0': { bold: true } } },
    ]);
    expect(new ExcelReader().parseFromBuffer(buffer).sheets[0].styles?.['0-0']).toMatchObject({
      bold: true,
    });
  });

  it.each([
    [20, 'h:mm'],
    [21, 'h:mm:ss'],
    [15, 'd-mmm-yy'],
    [22, 'm/d/yy h:mm'],
    [46, '[h]:mm:ss'],
  ])('reads built-in id %i as %s and writes it back', (id, code) => {
    const files = unzipSync(dateFile());
    files['xl/styles.xml'] = strToU8(
      strFromU8(files['xl/styles.xml']).replace('numFmtId="14"', `numFmtId="${id}"`)
    );
    const source = zipSync(files);
    expect(new ExcelReader().parseFromBuffer(source).sheets[0].styles?.['0-1']).toMatchObject({
      numberFormat: code,
    });
    expect(stylesXml(Workbook.fromBuffer(source).toBuffer())).toContain(`formatCode="${code}"`);
  });
});

const rewriteSheet = (transform: (xml: string) => string): Uint8Array => {
  const files = unzipSync(writer.createWorkbookBuffer([{ data: [['a', 'b', 1]] }]));
  files['xl/worksheets/sheet1.xml'] = strToU8(
    transform(strFromU8(files['xl/worksheets/sheet1.xml']))
  );
  return zipSync(files);
};

describe('loaded literals', () => {
  it('returns the loaded strings and saves error cells and = text as literals', () => {
    const source = rewriteSheet(xml =>
      xml
        .replace('<c r="A1" t="inlineStr"><is><t>a</t></is></c>', '<c r="A1" t="e"><v>#N/A</v></c>')
        .replace(
          '<c r="B1" t="inlineStr"><is><t>b</t></is></c>',
          '<c r="B1" t="inlineStr"><is><t>=== x ===</t></is></c>'
        )
    );
    const workbook = Workbook.fromBuffer(source);
    expect(workbook.getSheetData('Sheet1')[0]).toEqual(['#N/A', '=== x ===', 1]);
    const xml = part(workbook.toBuffer(), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" t="e"><v>#N/A</v></c>');
    expect(xml).toContain('<t>=== x ===</t>');
    workbook.setCellValue('Sheet1', 0, 0, '#N/A');
    expect(part(workbook.toBuffer(), 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="A1" t="inlineStr"><is><t>#N/A</t></is></c>'
    );
  });

  it('reads a blank numeric value as an empty cell and does not turn it into #NUM!', () => {
    const source = rewriteSheet(xml =>
      xml.replace('<c r="C1"><v>1</v></c>', '<c r="C1"><v></v></c>')
    );
    const cell = new ExcelReader().parseFromBuffer(source).sheets[0].data[0][2];
    expect(cell.type).toBe('empty');
    expect(part(Workbook.fromBuffer(source).toBuffer(), 'xl/worksheets/sheet1.xml')).not.toContain(
      '#NUM!'
    );
  });

  it('rejects an empty formula object but keeps the = shorthand bytes', () => {
    expect(() => writer.createWorkbookBuffer([{ data: [[{ formula: '' }]] }])).toThrow(
      'Cell A1 needs a formula'
    );
    expect(() => writer.createWorkbookBuffer([{ data: [['=']] }])).not.toThrow();
  });
});

describe('borders edge cases', () => {
  it('treats an object with no sides as no border', () => {
    const xml = stylesXml(
      writer.createWorkbookBuffer([{ data: [['a']], styles: { '0-0': { border: {} } } }])
    );
    expect(xml).toContain('<borders count="1">');
  });
});

describe('autofilter and layout', () => {
  it('shows the rows under the filter when the filter is removed', () => {
    const source = writer.createWorkbookBuffer([
      {
        data: [['h'], ['a'], ['b'], ['c']],
        options: { autoFilter: { range: 'A1:A4' }, hiddenRows: [1, 3] } as any,
      },
    ]);
    const workbook = Workbook.fromBuffer(source);
    expect(workbook.isRowHidden('Sheet1', 1)).toBe(true);
    workbook.removeAutoFilter('Sheet1');
    expect(workbook.isRowHidden('Sheet1', 1)).toBe(false);
    expect(workbook.isRowHidden('Sheet1', 3)).toBe(false);
  });

  it('survives being used as an array callback', () => {
    expect(() => [[10, 20], [5]].map(generateColsXml)).not.toThrow();
    expect([[10, 20]].map(generateColsXml)[0]).toBe(generateColsXml([10, 20]));
  });
});
