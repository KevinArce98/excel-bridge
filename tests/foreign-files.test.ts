import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { ExcelBridge, Workbook } from '../src';
import { knownDefect } from './helpers/known-defect';
import { calendarDay, cellAt, dateOf, fixture, normalise, part } from './helpers/read';

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
      expect(calendarDay(dateOf(cell))).toEqual([2024, 2, 29]);
    });
  }

  it('keeps every row where it was when the workbook is saved', () => {
    const [sheet] = ExcelBridge.read(Workbook.fromBuffer(bytes).toBuffer()).sheets;
    for (const [coordinate, value] of [
      ['A1', 'Region'],
      ['A3', 'North'],
      ['A5', 'Total'],
      ['A7', 'Footnote'],
    ]) {
      expect(cellAt(sheet, coordinate)?.value).toBe(value);
    }
    expect(sheet.mergeCells).toEqual(['A7:C7']);
  });

  it('keeps the hidden sheet hidden when the workbook is saved', () => {
    const workbook = Workbook.fromBuffer(bytes);
    expect(workbook.getSheetState('Hidden')).toBe('hidden');
    expect(workbook.getSheetState('Report')).toBe('visible');
    expect(part(workbook.toBuffer(), 'xl/workbook.xml')).toMatch(
      /name="Hidden"[^>]*state="hidden"/
    );
  });
});

describe('exceljs-report.xlsx validation', () => {
  const bytes = fixture('exceljs-report.xlsx');

  it('reads the whole-number rule on B3 in full', () => {
    expect(ExcelBridge.read(bytes).sheets[0].validations).toEqual([
      {
        range: 'B3',
        type: 'whole',
        formula1: '1',
        formula2: '10000',
        allowBlank: true,
      },
    ]);
  });

  it('keeps the rule as a whole-number rule when the workbook is saved', () => {
    const saved = Workbook.fromBuffer(bytes).toBuffer();
    expect(part(saved, 'xl/worksheets/sheet1.xml')).toMatch(
      /<dataValidation type="whole" operator="between" allowBlank="1"[^>]*sqref="B3">\s*<formula1>1<\/formula1><formula2>10000<\/formula2>/
    );
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
      expect(calendarDay(dateOf(cell))).toEqual([2024, 2, 29]);
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
