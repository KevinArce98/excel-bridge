import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { ExcelBridge, Workbook } from '../src';
import { knownDefect } from './helpers/known-defect';
import { calendarDay, cellAt, fixture, normalise, part } from './helpers/read';

const REPORTS = [
  {
    file: 'exceljs-report.xlsx',
    freezePane: { row: 1 },
    isoDates: false,
    headerFill: '#FFFF00',
    creator: 'excel-bridge fixtures',
  },
  {
    file: 'sheetjs-report.xlsx',
    freezePane: undefined,
    isoDates: true,
    headerFill: undefined,
    creator: undefined,
  },
  {
    file: 'hucre-report.xlsx',
    freezePane: { row: 1 },
    isoDates: false,
    headerFill: '#FFFF00',
    creator: undefined,
  },
];

describe.each(REPORTS)('$file', ({ file, freezePane, isoDates, headerFill, creator }) => {
  const bytes = fixture(file);
  const workbook = ExcelBridge.read(bytes);
  const report = workbook.sheets[0];

  it('matches the recorded parse of the file', () => {
    expect(normalise(workbook)).toMatchSnapshot();
  });

  it('lists both sheets in order', () => {
    expect(workbook.sheets.map(sheet => sheet.name)).toEqual(['Report', 'Hidden']);
    expect(cellAt(workbook.sheets[1], 'A1')?.value).toBe('secret');
  });

  it('reads text and numbers at their coordinates', () => {
    expect(cellAt(report, 'A1')?.value).toBe('Region');
    expect(cellAt(report, 'B1')?.value).toBe('Revenue');
    expect(cellAt(report, 'A3')?.value).toBe('North');
    expect(cellAt(report, 'B3')?.value).toBe(1200.5);
    expect(cellAt(report, 'A5')?.value).toBe('Total');
    expect(cellAt(report, 'A7')?.value).toBe('Footnote');
  });

  it('reads the formula text', () => {
    expect(cellAt(report, 'B5')?.formula).toBe('SUM(B1:B4)');
  });

  it('reads the merged range, the freeze pane and the column widths', () => {
    expect(report.mergeCells).toEqual(['A7:C7']);
    expect(report.freezePane).toEqual(freezePane);
    expect(report.columnWidths?.slice(0, 3).map(Math.floor)).toEqual([14, 12, 12]);
  });

  it('reads the header fill', () => {
    expect(report.styles?.['0-0']?.background).toBe(headerFill);
  });

  it('reads the document creator', () => {
    expect(workbook.metadata.creator).toBe(creator);
  });

  it('is stored with the second sheet marked hidden', () => {
    expect(part(bytes, 'xl/workbook.xml')).toMatch(/state="hidden"/);
  });

  if (!isoDates) {
    it('reads the date cell as a calendar day', () => {
      const cell = cellAt(report, 'C3');
      expect(cell?.type).toBe('date');
      expect(calendarDay(cell?.value)).toEqual([2024, 2, 29]);
    });
  }

  knownDefect(
    'rows after a blank row move up when the workbook is saved (expected: A5 stays "Total") (I1)',
    () => {
      const [sheet] = ExcelBridge.read(Workbook.fromBuffer(bytes).toBuffer()).sheets;
      expect(cellAt(sheet, 'A5')?.value).toBe('Total');
    },
    { message: /expected undefined to be 'Total'/ }
  );

  knownDefect(
    'a hidden sheet becomes visible when the workbook is saved (expected: state="hidden") (I6)',
    () => {
      expect(part(Workbook.fromBuffer(bytes).toBuffer(), 'xl/workbook.xml')).toMatch(
        /state="hidden"/
      );
    },
    { message: /to match/ }
  );
});

describe('exceljs-report.xlsx validation', () => {
  it('stores a whole-number rule on B3', () => {
    expect(ExcelBridge.read(fixture('exceljs-report.xlsx')).sheets[0].validations).toEqual([
      { range: 'B3', options: '1' },
    ]);
  });
});

describe('sheetjs-report.xlsx stores its date as ISO text', () => {
  const bytes = fixture('sheetjs-report.xlsx');

  it('marks the cell as t="d" with the exact instant', () => {
    expect(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'])).toMatch(
      /<c r="C3"[^>]*t="d"[^>]*><v>2024-02-29T00:00:00\.000Z<\/v>/
    );
  });

  knownDefect(
    'an ISO-8601 date cell is read as year 1905 (expected 2024-02-29) (R7)',
    () => {
      const cell = cellAt(ExcelBridge.read(bytes).sheets[0], 'C3');
      expect(calendarDay(cell?.value)).toEqual([2024, 2, 29]);
    },
    { message: /expected \[ 1905, 7, 16 \] to deeply equal \[ 2024, 2, 29 \]/ }
  );
});

describe('exceljs-text.xlsx', () => {
  const sheet = ExcelBridge.read(fixture('exceljs-text.xlsx')).sheets[0];

  it('matches the recorded parse of the file', () => {
    expect(normalise(ExcelBridge.read(fixture('exceljs-text.xlsx')))).toMatchSnapshot();
  });

  it.each([
    ['A1', 'bold  plain'],
    ['A2', '  padded on both sides  '],
    ['A3', `& < > " '`],
    ['A4', 'line one\nline two'],
    ['A5', '😀 日本語 مرحبا'],
    ['A6', '00123'],
    ['A7', 'x'.repeat(300)],
    ['A8', '&#233; &amp;'],
  ])('reads %s exactly', (coordinate, expected) => {
    expect(cellAt(sheet, coordinate)?.value).toBe(expected);
    expect(cellAt(sheet, coordinate)?.type).toBe('string');
  });
});
