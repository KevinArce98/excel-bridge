import { describe, expect, it } from 'vitest';
import { ExcelBridge, Workbook } from '../src';
import { rowIndexes } from '../src/core/rows';
import { cellAt } from './helpers/read';

const sparse = (length: number, populated: number[]): string[][] => {
  const rows: string[][] = [];
  rows.length = length;
  populated.forEach(index => {
    rows[index] = [`row ${index}`];
  });
  return rows;
};

describe.each([
  ['a short sparse list', 100, [3, 40, 99]],
  ['a long sparse list', 1_000_000, [0, 70_000, 948_575, 999_999]],
])('rowIndexes on %s', (_label, length, populated) => {
  it('lists only the populated rows, in order', () => {
    expect(rowIndexes(sparse(length, populated))).toEqual(populated);
  });
});

describe('rowIndexes on dense lists', () => {
  it.each([0, 1, 10, 70_000])('lists every row of %i rows', length => {
    const rows = Array.from({ length }, (_, index) => [index]);
    expect(rowIndexes(rows)).toEqual(Array.from({ length }, (_, index) => index));
  });
});

describe('saving workbooks whose rows are far apart', () => {
  it('does not scan the empty rows of every sheet', () => {
    const workbook = Workbook.create();
    for (let index = 0; index < 200; index++) {
      workbook.addSheet(`S${index}`, [['a']]);
      workbook.setCellValue(`S${index}`, 1_048_575, 0, 'last');
    }

    const started = performance.now();
    const bytes = workbook.toBuffer();
    const elapsed = performance.now() - started;

    expect(elapsed).toBeLessThan(2000);
    const sheets = ExcelBridge.read(bytes).sheets;
    expect(sheets).toHaveLength(200);
    expect(cellAt(sheets[199], 'A1048576')?.value).toBe('last');
  });
});
