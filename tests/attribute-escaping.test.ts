import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { XMLValidator } from 'fast-xml-parser';
import { ExcelWriter, ExcelBridge } from '../src';
import type { CellValidation, ConditionalFormat } from '../src';

const hostile = 'A1"/><x a="&';

describe('Attribute escaping', () => {
  it('round-trips a sheet name with quotes and markup', () => {
    const name = 'Q1 "Sales" & <Co>';
    const buffer = new ExcelWriter().createWorkbookBuffer([{ data: [['x']], options: { name } }]);

    expect(strFromU8(unzipSync(buffer)['xl/workbook.xml'])).toContain(
      'name="Q1 &quot;Sales&quot; &amp; &lt;Co&gt;"'
    );
    expect(ExcelBridge.read(buffer).sheets[0].name).toBe(name);
  });

  it('round-trips a font name with quotes', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: [['x']], styles: { '0-0': { fontName: 'Fira "Code" & Co' } } },
    ]);

    expect(ExcelBridge.read(buffer).sheets[0].styles?.['0-0']).toMatchObject({
      fontName: 'Fira "Code" & Co',
    });
  });

  it('keeps worksheet and style XML well-formed whatever the input', () => {
    const conditionalFormats = [
      { type: 'expression', range: hostile, formula: 'TRUE', style: { bold: true } },
      {
        type: 'cellValue',
        range: hostile,
        operator: hostile,
        value: 1,
        style: { color: '#9C0006', background: '#FFC7CE' },
      },
      { type: 'colorScale', range: hostile, colors: ['#FF0000', '#FFFFFF'] },
    ] as unknown as ConditionalFormat[];
    const validations = [
      { range: hostile, type: 'list', options: '', formula1: '"a&b<c"' },
      { range: hostile, type: hostile, operator: hostile, formula1: '1', options: '' },
    ] as unknown as CellValidation[];

    const buffer = new ExcelWriter().createWorkbookBuffer([
      {
        data: [['x']],
        mergeCells: [hostile],
        conditionalFormats,
        validations,
        styles: {
          '0-0': {
            align: hostile,
            verticalAlign: hostile,
          } as never,
        },
      },
      { data: [['y']], options: { name: 'Second', state: hostile as never } },
    ]);
    const files = unzipSync(buffer);

    for (const path of ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml']) {
      expect(XMLValidator.validate(strFromU8(files[path])), path).toBe(true);
    }
  });

  it('escapes markup in the sheet state', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: [['x']] },
      { data: [['y']], options: { name: 'Second', state: hostile as never } },
    ]);
    expect(strFromU8(unzipSync(buffer)['xl/workbook.xml'])).toContain(
      'state="A1&quot;/&gt;&lt;x a=&quot;&amp;"'
    );
  });

  it('rejects markup in a colour or a sheet name instead of escaping it', () => {
    const write = (sheet: Parameters<ExcelWriter['createWorkbookBuffer']>[0][number]) =>
      new ExcelWriter().createWorkbookBuffer([sheet]);

    expect(() => write({ data: [['x']], styles: { '0-0': { color: hostile } } })).toThrow(
      /Invalid colour/
    );
    expect(() => write({ data: [['x']], styles: { '0-0': { background: hostile } } })).toThrow(
      /Invalid colour/
    );
    expect(() => write({ data: [['x']], options: { name: hostile } })).toThrow(
      /Invalid sheet name/
    );
  });
});
