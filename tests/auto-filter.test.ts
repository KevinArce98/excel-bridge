import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8, zipSync, strToU8 } from 'fflate';
import { ExcelWriter, ExcelBridge, Workbook } from '../src';
import type { CellValue } from '../src';

const rows: CellValue[][] = [
  ['Region', 'Rep', 'Units', 'Revenue'],
  ['North', 'Ann', 10, 1200],
  ['South', 'Bob', 4, 480],
  ['East', 'Cy', 7, 910],
];

const part = (buffer: Uint8Array, path: string) => strFromU8(unzipSync(buffer)[path]);

describe('AutoFilter', () => {
  it('writes the autoFilter ref and normalizes the range', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: rows, options: { autoFilter: { range: '$d$4:a1' } } },
    ]);

    expect(part(buffer, 'xl/worksheets/sheet1.xml')).toContain('<autoFilter ref="A1:D4"/>');
  });

  it('adds a hidden _FilterDatabase name scoped to the filtered sheet', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: rows },
      { data: rows, options: { name: "Bob's Q1 & Co", autoFilter: { range: 'A1:D4' } } },
    ]);
    const workbook = part(buffer, 'xl/workbook.xml');

    expect(workbook).toContain(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'Bob''s Q1 &amp; Co'!$A$1:$D$4</definedName>`
    );
    expect(workbook.match(/<definedName /g)).toHaveLength(1);
  });

  it('writes one hidden name per filtered sheet', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: rows, options: { name: 'North', autoFilter: { range: 'A1:D4' } } },
      { data: rows, options: { name: 'South', autoFilter: { range: 'A1:B2' } } },
    ]);
    const workbook = part(buffer, 'xl/workbook.xml');

    expect(workbook).toContain(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'North'!$A$1:$D$4</definedName>`
    );
    expect(workbook).toContain(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'South'!$A$1:$B$2</definedName>`
    );
  });

  it('writes no definedNames when no sheet has a filter', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([{ data: rows }]);
    expect(part(buffer, 'xl/workbook.xml')).not.toContain('definedNames');
  });

  it.each([
    ['A0:B2', 'Invalid range format: A0:B2'],
    ['A1:B2:C3', 'Invalid range format: A1:B2:C3'],
    ['Sales', 'Invalid range format: Sales'],
    ['XFE1', 'Column index 16384 exceeds Excel limit (0-16383)'],
    ['A1048577', 'Row index 1048576 exceeds Excel limit (0-1048575)'],
  ])('rejects %s', (range, message) => {
    expect(() =>
      new ExcelWriter().createWorkbookBuffer([{ data: rows, options: { autoFilter: { range } } }])
    ).toThrow(message);
  });

  it('reads the range back and omits the field when absent', () => {
    const filtered = ExcelBridge.read(
      new ExcelWriter().createWorkbookBuffer([
        { data: rows, options: { autoFilter: { range: 'A1:D4' } } },
      ])
    );
    expect(filtered.sheets[0].autoFilter).toEqual({ range: 'A1:D4' });

    const plain = ExcelBridge.read(new ExcelWriter().createWorkbookBuffer([{ data: rows }]));
    expect(plain.sheets[0]).not.toHaveProperty('autoFilter');
  });

  it('reads a filter saved with criteria as its range', () => {
    const files = unzipSync(new ExcelWriter().createWorkbookBuffer([{ data: rows }]));
    const sheet = strFromU8(files['xl/worksheets/sheet1.xml']).replace(
      '</sheetData>',
      '</sheetData><autoFilter ref="A1:D4"><filterColumn colId="0"><filters><filter val="North"/></filters></filterColumn></autoFilter>'
    );
    files['xl/worksheets/sheet1.xml'] = strToU8(sheet);

    expect(ExcelBridge.read(zipSync(files)).sheets[0].autoFilter).toEqual({ range: 'A1:D4' });
  });

  describe('Workbook', () => {
    it('keeps the filter through load, edit and save', () => {
      const buffer = new ExcelWriter().createWorkbookBuffer([
        { data: rows, options: { name: 'Sales', autoFilter: { range: 'A1:D4' } } },
      ]);
      const workbook = Workbook.fromBuffer(buffer);
      workbook.setCellValue('Sales', 1, 2, 11);

      const saved = workbook.toBuffer();
      expect(ExcelBridge.read(saved).sheets[0].autoFilter).toEqual({ range: 'A1:D4' });
      expect(part(saved, 'xl/workbook.xml')).toContain('_xlnm._FilterDatabase');
    });

    it('sets, gets and removes the filter', () => {
      const workbook = Workbook.create();
      workbook.addSheet('Sales', rows);

      workbook.setAutoFilter('Sales', { range: 'd4:A1' });
      expect(workbook.getAutoFilter('Sales')).toEqual({ range: 'A1:D4' });

      workbook.removeAutoFilter('Sales');
      expect(workbook.getAutoFilter('Sales')).toBeUndefined();
      expect(part(workbook.toBuffer(), 'xl/worksheets/sheet1.xml')).not.toContain('autoFilter');
    });

    it('validates at the call', () => {
      const workbook = Workbook.create();
      workbook.addSheet('Sales', rows);

      expect(() => workbook.setAutoFilter('Sales', { range: 'A1:' })).toThrow(
        'Invalid range format: A1:'
      );
      expect(() => workbook.setAutoFilter('Missing', { range: 'A1:D4' })).toThrow(
        'Sheet "Missing" not found'
      );
    });
  });
});
