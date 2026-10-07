import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { ExcelBridge, Workbook } from '../src';
import type { ParsedSheet } from '../src/reader';
import { knownDefect } from './helpers/known-defect';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));

const cellAt = (sheet: ParsedSheet, coordinate: string) =>
  sheet.data.flat().find(cell => cell?.coordinate === coordinate);

const calendarDay = (date: Date) => [date.getFullYear(), date.getMonth() + 1, date.getDate()];

const workbookXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/workbook.xml']);

const FIXTURES = [
  { file: 'exceljs-report.xlsx', freezePane: { row: 1 }, isoDates: false },
  { file: 'sheetjs-report.xlsx', freezePane: undefined, isoDates: true },
  { file: 'hucre-report.xlsx', freezePane: { row: 1 }, isoDates: false },
];

describe.each(FIXTURES)('$file', ({ file, freezePane, isoDates }) => {
  const bytes = fixture(file);
  const workbook = ExcelBridge.read(bytes);
  const report = workbook.sheets[0];

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

  it('reads the merged range and the freeze pane', () => {
    expect(report.mergeCells).toEqual(['A7:C7']);
    expect(report.freezePane).toEqual(freezePane);
  });

  it('keeps the hidden state in the source file', () => {
    expect(workbookXml(bytes)).toMatch(/state="hidden"/);
  });

  if (!isoDates) {
    it('reads the date cell as a calendar day', () => {
      const cell = cellAt(report, 'C3');
      expect(cell?.type).toBe('date');
      expect(calendarDay(cell?.value)).toEqual([2024, 2, 29]);
    });
  }

  knownDefect('I1: Workbook round trip keeps rows where they were', () => {
    const [sheet] = ExcelBridge.read(Workbook.fromBuffer(bytes).toBuffer()).sheets;
    expect(cellAt(sheet, 'A5')?.value).toBe('Total');
  });

  knownDefect('I6: Workbook round trip keeps a hidden sheet hidden', () => {
    expect(workbookXml(Workbook.fromBuffer(bytes).toBuffer())).toMatch(/state="hidden"/);
  });
});

describe('sheetjs-report.xlsx stores its date as ISO text', () => {
  const bytes = fixture('sheetjs-report.xlsx');

  it('marks the cell as t="d"', () => {
    expect(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'])).toMatch(/<c r="C3"[^>]*t="d"/);
  });

  knownDefect('R7: an ISO-8601 date cell is read as a calendar day', () => {
    const cell = cellAt(ExcelBridge.read(bytes).sheets[0], 'C3');
    expect(calendarDay(cell?.value)).toEqual([2024, 2, 29]);
  });
});
