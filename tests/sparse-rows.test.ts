import { describe, expect, it } from 'vitest';
import { ExcelBridge, ExcelReader, Workbook } from '../src';
import { rowIndexes } from '../src/core/rows';
import { cellAt, part, readFirstSheet } from './helpers/read';
import { buildXlsx } from './helpers/xlsx';

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

describe('ParsedSheet.data is indexed by row index', () => {
  const bytes = buildXlsx(
    '<sheetData><row r="5"><c r="A5"><v>5</v></c></row><row r="2"><c r="B2"><v>1</v></c></row><row r="9" ht="20" customHeight="1"/><row r="12"/></sheetData>'
  );
  const sheet = readFirstSheet(bytes);

  it('puts each row at its index, in index order, and leaves holes elsewhere', () => {
    expect(Object.keys(sheet.data)).toEqual(['1', '4']);
    expect(sheet.data).toHaveLength(5);
    expect(sheet.data[1][0].rowIndex).toBe(1);
    expect(sheet.data[4][0].coordinate).toBe('A5');
  });

  it('pads a row with empty cells from column A to its last cell', () => {
    expect(sheet.data[1].map(cell => cell.type)).toEqual(['empty', 'number']);
  });

  it('leaves a row element without cells as a hole and keeps its layout', () => {
    expect(sheet.data[8]).toBeUndefined();
    expect(sheet.data[11]).toBeUndefined();
    expect(sheet.rowHeights).toEqual({ 8: 20 });
  });

  it('lists the rows that are present with forEach, Object.values and flat', () => {
    const seen: number[] = [];
    sheet.data.forEach((_row, rowIndex) => seen.push(rowIndex));
    expect(seen).toEqual([1, 4]);
    expect(Object.values(sheet.data)).toHaveLength(2);
    expect(sheet.data.flat().map(cell => cell.coordinate)).toEqual(['A2', 'B2', 'A5']);
  });

  it('gives Workbook the same holes', () => {
    const data = Workbook.fromBuffer(bytes).getSheetData('S');
    expect(Object.keys(data)).toEqual(['1', '4']);
    expect(data).toHaveLength(5);
  });

  it('does not add empty rows when a layout-only row is saved', () => {
    const saved = part(Workbook.fromBuffer(bytes).toBuffer(), 'xl/worksheets/sheet1.xml');
    expect(saved.match(/<row r="\d+"/g)).toEqual(['<row r="2"', '<row r="5"', '<row r="9"']);
  });

  it('reads a row at the end of the grid without allocating the rows before it', () => {
    const far = buildXlsx(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="1048576"><c r="A1048576"><v>2</v></c></row></sheetData>'
    );
    const started = performance.now();
    const data = readFirstSheet(far).data;
    expect(performance.now() - started).toBeLessThan(1000);
    expect(data).toHaveLength(1048576);
    expect(Object.keys(data)).toEqual(['0', '1048575']);
  });

  it.each(['r="0"', 'r="x"', 'r=""'])('treats %s like a missing r attribute', attribute => {
    const sheet = readFirstSheet(
      buildXlsx(
        `<sheetData><row r="3"><c r="A3"><v>1</v></c></row><row ${attribute}><c><v>2</v></c></row></sheetData>`
      )
    );
    expect(sheet.data[3][0].coordinate).toBe('A4');
  });
});

describe('rows and cells that are not where they say', () => {
  it('gives a cell with no r the column after the previous cell, in the middle of a row', () => {
    const sheet = readFirstSheet(
      buildXlsx(
        '<sheetData><row r="1"><c r="B1"><v>1</v></c><c><v>2</v></c><c r="F1"><v>3</v></c><c><v>4</v></c></row></sheetData>'
      )
    );
    expect(sheet.data[0].map(cell => cell.coordinate)).toEqual([
      'A1',
      'B1',
      'C1',
      'D1',
      'E1',
      'F1',
      'G1',
    ]);
    expect(sheet.data[0].map(cell => cell.value)).toEqual([null, 1, 2, null, null, 3, 4]);
    expect(sheet.data[0].every(cell => Number.isInteger(cell.columnIndex))).toBe(true);
  });

  it('merges two row elements with the same r into one row, the later cell winning', () => {
    const sheet = readFirstSheet(
      buildXlsx(
        '<sheetData><row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>2</v></c></row><row r="3"><c r="B3"><v>9</v></c><c r="D3"><v>4</v></c></row></sheetData>'
      )
    );
    expect(Object.keys(sheet.data)).toEqual(['2']);
    expect(sheet.data[2].map(cell => cell.value)).toEqual([1, 9, null, 4]);
  });

  it('counts the cells of a merged row once against maxCells', () => {
    const body =
      '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="1"><c r="A1"><v>3</v></c><c r="B1"><v>4</v></c></row></sheetData>';
    expect(() => new ExcelReader({ maxCells: 2 }).parseFromBuffer(buildXlsx(body))).not.toThrow();
    expect(() => new ExcelReader({ maxCells: 1 }).parseFromBuffer(buildXlsx(body))).toThrow(
      /cells/
    );
  });
});
