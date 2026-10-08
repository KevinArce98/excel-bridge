import { describe, expect, it } from 'vitest';
import {
  ExcelBridge,
  ExcelWriter,
  Workbook,
  createExcelWorkbookStream,
  streamToBuffer,
} from '../src';
import { generateColsXml } from '../src/core/column-width';
import type { ExcelData, SheetOptions } from '../src';
import { cellAt, part, readFirstSheet } from './helpers/read';
import { buildXlsx } from './helpers/xlsx';

const sheetXml = (bytes: Uint8Array) => part(bytes, 'xl/worksheets/sheet1.xml');
const writeSheet = (data: ExcelData['data'], options: SheetOptions) =>
  new ExcelWriter().createWorkbookBuffer([{ data, options }]);
const rowTags = (xml: string) => [...xml.matchAll(/<row r="(\d+)"([^>]*)>/g)].map(m => m.slice(1));

const grid = [['a'], ['b'], ['c']];

describe('row heights and hidden rows or columns in ExcelWriter', () => {
  it('writes a manual height with customHeight, and a hidden row with hidden', () => {
    const rows = rowTags(
      sheetXml(writeSheet(grid, { rowHeights: { 0: 30, 2: 22.5 }, hiddenRows: [1, 2] }))
    );
    expect(rows).toEqual([
      ['1', ' ht="30" customHeight="1"'],
      ['2', ' hidden="1"'],
      ['3', ' ht="22.5" customHeight="1" hidden="1"'],
    ]);
  });

  it('leaves rows without layout exactly as before', () => {
    expect(rowTags(sheetXml(writeSheet(grid, {})))).toEqual([
      ['1', ''],
      ['2', ''],
      ['3', ''],
    ]);
  });

  it('writes rows that only carry layout, in ascending order between and after data rows', () => {
    const sparse: string[][] = [];
    sparse[0] = ['first'];
    sparse[10] = ['eleventh'];
    const xml = sheetXml(
      writeSheet(sparse, { rowHeights: { 5: 40, 20: 12 }, hiddenRows: [3, 10, 15] })
    );
    const indexes = rowTags(xml).map(([r]) => Number(r));
    expect(indexes).toEqual([1, 4, 6, 11, 16, 21]);
    expect(xml).toContain('<row r="6" ht="40" customHeight="1"></row>');
  });

  it('accepts an array of heights', () => {
    const rows = rowTags(sheetXml(writeSheet(grid, { rowHeights: [30, 20] })));
    expect(rows[0][1]).toBe(' ht="30" customHeight="1"');
    expect(rows[1][1]).toBe(' ht="20" customHeight="1"');
  });

  it.each([
    [{ rowHeights: { 0: 0 } }, 'Row 1 height 0 must be above 0 and at most 409.5 points'],
    [{ rowHeights: { 0: -5 } }, 'Row 1 height -5'],
    [{ rowHeights: { 1: NaN } }, 'Row 2 height NaN'],
    [{ rowHeights: { 1: Infinity } }, 'Row 2 height Infinity'],
    [{ rowHeights: { 0: 409.6 } }, 'Row 1 height 409.6'],
    [{ rowHeights: { 0: '30' as unknown as number } }, 'Row 1 height 30'],
    [{ rowHeights: { 1.5: 30 } }, 'Row index 1.5 must be a whole number from 0 to 1048575'],
    [{ rowHeights: { '-1': 30 } }, 'Row index -1'],
    [{ rowHeights: { 1048576: 30 } }, 'Row index 1048576'],
    [{ rowHeights: { abc: 30 } }, 'Row index NaN'],
    [{ hiddenRows: [1.5] }, 'Row index 1.5'],
    [{ hiddenRows: [NaN] }, 'Row index NaN'],
    [{ hiddenRows: [1048576] }, 'Row index 1048576'],
    [{ hiddenColumns: [16384] }, 'Column index 16384 must be a whole number from 0 to 16383'],
    [{ hiddenColumns: [-1] }, 'Column index -1'],
  ])('rejects %j', (options, message) => {
    expect(() => writeSheet(grid, options as SheetOptions)).toThrow(message);
  });

  it('accepts Excel limits', () => {
    expect(() =>
      writeSheet(grid, {
        rowHeights: { 0: 409.5, 1: Number.MIN_VALUE, 1048575: 1 },
        hiddenRows: [1048575],
        hiddenColumns: [16383],
      })
    ).not.toThrow();
  });

  it('writes hidden columns, keeping the declared width and defaulting the rest', () => {
    const xml = sheetXml(writeSheet(grid, { columnWidths: [12, 10], hiddenColumns: [1, 3] }));
    expect(xml).toContain('<col min="1" max="1" width="12" customWidth="1"/>');
    expect(xml).toContain('<col min="2" max="2" width="10" customWidth="1" hidden="1"/>');
    expect(xml).toContain('<col min="4" max="4" width="9.140625" customWidth="1" hidden="1"/>');
  });

  it('hides a column next to the widths autoWidth measured', () => {
    const xml = sheetXml(
      writeSheet([['a long enough text', 'b']], { autoWidth: true, hiddenColumns: [1] })
    );
    expect(xml).toMatch(/<col min="1" max="1" width="\d+" customWidth="1"\/>/);
    expect(xml).toMatch(/<col min="2" max="2" width="\d+" customWidth="1" hidden="1"\/>/);
  });

  it('writes the same <cols> as before when no column is hidden', () => {
    const sparseWidths: number[] = [];
    sparseWidths[2] = 12;
    expect(generateColsXml([10, 20])).toBe(
      '  <cols>\n    <col min="1" max="1" width="10" customWidth="1"/>\n    <col min="2" max="2" width="20" customWidth="1"/>\n  </cols>'
    );
    expect(generateColsXml(sparseWidths)).toBe(
      '  <cols>\n\n\n    <col min="3" max="3" width="12" customWidth="1"/>\n  </cols>'
    );
    expect(generateColsXml([])).toBe('');
  });

  it('does not scan the empty rows of sheets whose rows are far apart', () => {
    const rows: string[][] = [];
    rows[0] = ['a'];
    rows[1_048_575] = ['last'];
    const started = performance.now();
    for (let index = 0; index < 50; index++) writeSheet(rows, { rowHeights: { 0: 30 } });
    expect(performance.now() - started).toBeLessThan(2000);
  });
});

describe('the streaming writer', () => {
  const options: SheetOptions = {
    rowHeights: { 0: 30, 2: 22.5, 9: 40 },
    hiddenRows: [1, 12],
    hiddenColumns: [1],
    columnWidths: [12, 10],
  };

  it('writes the same sheet as ExcelWriter, including rows past the last one streamed', async () => {
    const streamed = await streamToBuffer(createExcelWorkbookStream([{ rows: grid, ...options }]));
    expect(sheetXml(streamed)).toBe(sheetXml(writeSheet(grid, options)));
    expect(rowTags(sheetXml(streamed)).map(([r]) => r)).toEqual(['1', '2', '3', '10', '13']);
  });

  it('decides layout from the row index for an async iterable', async () => {
    async function* rows() {
      for (let index = 0; index < 5; index++) yield [index];
    }
    const buffer = await streamToBuffer(
      createExcelWorkbookStream([{ rows: rows(), rowHeights: { 3: 18 } }])
    );
    expect(rowTags(sheetXml(buffer))[3]).toEqual(['4', ' ht="18" customHeight="1"']);
  });

  it('rejects invalid layout before yielding anything', async () => {
    const stream = createExcelWorkbookStream([{ rows: grid, rowHeights: { 0: 500 } }]);
    await expect(stream.next()).rejects.toThrow('Row 1 height 500');
  });
});

describe('ExcelReader', () => {
  const read = (body: string) => readFirstSheet(buildXlsx(body));
  const row = (attributes: string) =>
    `<sheetData><row r="2" ${attributes}><c r="A2"><v>1</v></c></row></sheetData>`;

  it.each([
    ['Excel and ExcelJS', 'ht="30" customHeight="1" spans="1:1"'],
    ['SheetJS', 'hidden="0" ht="30" customHeight="true"'],
  ])('reads a manual height as %s writes it', (_tool, attributes) => {
    expect(read(row(attributes)).rowHeights).toEqual({ 1: 30 });
  });

  it.each([
    ['an automatic height', 'ht="30"'],
    ['customHeight 0', 'ht="30" customHeight="0"'],
    ['customHeight false', 'ht="30" customHeight="false"'],
    ['a missing ht', 'customHeight="1"'],
    ['a zero height', 'ht="0" customHeight="1"'],
    ['a negative height', 'ht="-3" customHeight="1"'],
    ['a height past 409.5', 'ht="410" customHeight="1"'],
    ['text', 'ht="tall" customHeight="1"'],
  ])('ignores %s', (_label, attributes) => {
    expect(read(row(attributes)).rowHeights).toBeUndefined();
  });

  it.each(['1', 'true'])('reads hidden="%s"', value => {
    expect(read(row(`hidden="${value}"`)).hiddenRows).toEqual([1]);
  });

  it.each(['0', 'false'])('does not read hidden="%s" as hidden', value => {
    expect(read(row(`hidden="${value}"`)).hiddenRows).toBeUndefined();
  });

  it('keeps an empty row that carries only layout', () => {
    const sheet = read('<sheetData><row r="3" ht="40" customHeight="1" hidden="1"/></sheetData>');
    expect(sheet.rowHeights).toEqual({ 2: 40 });
    expect(sheet.hiddenRows).toEqual([2]);
  });

  it('applies the layout of a row with no r attribute to the row after the previous one', () => {
    const sheet = read(
      '<sheetData><row r="2"><c r="A2"><v>1</v></c></row><row ht="30" customHeight="1" hidden="1"><c r="A3"><v>1</v></c></row></sheetData>'
    );
    expect(sheet.rowHeights).toEqual({ 2: 30 });
    expect(sheet.hiddenRows).toEqual([2]);
  });

  it('numbers a first row with no r attribute as row 1', () => {
    const sheet = read(
      '<sheetData><row ht="30" customHeight="1"><c><v>1</v></c><c><v>2</v></c></row></sheetData>'
    );
    expect(sheet.rowHeights).toEqual({ 0: 30 });
    expect(sheet.data[0].map(cell => cell.coordinate)).toEqual(['A1', 'B1']);
  });

  it('reads hidden columns from ranges, with true or 1, and keeps their width', () => {
    const sheet = read(
      '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="3" width="10" hidden="true" customWidth="1"/><col min="5" max="5" width="9" hidden="1"/></cols><sheetData/>'
    );
    expect(sheet.hiddenColumns).toEqual([1, 2, 4]);
    expect(sheet.columnWidths).toEqual([12, 10, 10, undefined, 9]);
  });

  it('bounds a hidden column range to the grid', () => {
    const sheet = read(
      '<cols><col min="1" max="100000000" width="10" hidden="1"/></cols><sheetData/>'
    );
    expect(sheet.hiddenColumns).toHaveLength(16384);
  });

  it('leaves a hole instead of NaN for a column without a width', () => {
    const sheet = read('<cols><col min="2" max="2" hidden="1" style="1"/></cols><sheetData/>');
    expect(sheet.columnWidths?.some(Number.isNaN)).toBe(false);
    expect(sheet.hiddenColumns).toEqual([1]);
  });
});

describe('a ExcelWriter -> reader -> Workbook round trip', () => {
  const options: SheetOptions = {
    name: 'Main',
    rowHeights: { 0: 30, 2: 22.5, 6: 40 },
    hiddenRows: [1, 8],
    hiddenColumns: [1, 3],
    columnWidths: [12, 10, 8, 9],
    freezePane: { row: 1 },
  };
  const original = writeSheet(grid, options);

  it('reads back what was declared', () => {
    const sheet = readFirstSheet(original);
    expect(sheet.rowHeights).toEqual(options.rowHeights);
    expect(sheet.hiddenRows).toEqual([1, 8]);
    expect(sheet.hiddenColumns).toEqual([1, 3]);
  });

  it('keeps the layout of a Workbook that is loaded and saved', () => {
    const sheet = readFirstSheet(Workbook.fromBuffer(original).toBuffer());
    expect(sheet.rowHeights).toEqual(options.rowHeights);
    expect(sheet.hiddenRows).toEqual(options.hiddenRows);
    expect(sheet.hiddenColumns).toEqual(options.hiddenColumns);
  });

  it('writes the sheet unchanged when every layout row has data', () => {
    const dense = writeSheet(grid, {
      name: 'Main',
      rowHeights: { 0: 30, 2: 22.5 },
      hiddenRows: [1],
      hiddenColumns: [1],
      columnWidths: [12, 10],
    });
    expect(sheetXml(Workbook.fromBuffer(dense).toBuffer())).toBe(sheetXml(dense));
  });

  it('edits heights and visibility on a loaded workbook', () => {
    const workbook = Workbook.fromBuffer(original);
    expect(workbook.getRowHeight('Main', 0)).toBe(30);
    expect(workbook.getRowHeight('Main', 1)).toBeUndefined();
    expect(workbook.isRowHidden('Main', 1)).toBe(true);
    expect(workbook.isColumnHidden('Main', 1)).toBe(true);

    workbook.setRowHeight('Main', 0, null);
    workbook.setRowHeight('Main', 4, 50);
    workbook.setRowHidden('Main', 1, false);
    workbook.setRowHidden('Main', 5);
    workbook.setColumnHidden('Main', 1, false);

    const sheet = readFirstSheet(workbook.toBuffer());
    expect(sheet.rowHeights).toEqual({ 2: 22.5, 4: 50, 6: 40 });
    expect(sheet.hiddenRows).toEqual([5, 8]);
    expect(sheet.hiddenColumns).toEqual([3]);
  });

  it('builds a workbook from scratch', () => {
    const workbook = Workbook.create();
    workbook.addSheet('S', grid);
    workbook.setRowHeight('S', 1, 33);
    workbook.setRowHidden('S', 2);
    workbook.setColumnHidden('S', 2);
    const sheet = readFirstSheet(workbook.toBuffer());
    expect(sheet.rowHeights).toEqual({ 1: 33 });
    expect(sheet.hiddenRows).toEqual([2]);
    expect(sheet.hiddenColumns).toEqual([2]);
  });

  it.each([
    [
      'a height above 409.5',
      (workbook: Workbook) => workbook.setRowHeight('S', 0, 500),
      'Row 1 height 500',
    ],
    ['a fractional row', (workbook: Workbook) => workbook.setRowHidden('S', 1.5), 'Row index 1.5'],
    [
      'a column past XFD',
      (workbook: Workbook) => workbook.setColumnHidden('S', 16384),
      'Column index 16384',
    ],
  ])('refuses %s when it is set', (_label, act, message) => {
    const workbook = Workbook.create();
    workbook.addSheet('S', grid);
    expect(() => act(workbook)).toThrow(message);
    expect(() => workbook.toBuffer()).not.toThrow();
  });

  it('keeps hidden rows on a sheet with a filter, without the filter criteria', () => {
    const body =
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row><row r="2" hidden="1"><c r="A2"><v>1</v></c></row></sheetData><autoFilter ref="A1:A2"><filterColumn colId="0"><filters><filter val="2"/></filters></filterColumn></autoFilter>';
    const saved = Workbook.fromBuffer(buildXlsx(body)).toBuffer();
    const sheet = readFirstSheet(saved);
    expect(sheet.autoFilter).toEqual({ range: 'A1:A2' });
    expect(sheet.hiddenRows).toEqual([1]);
    expect(sheetXml(saved)).not.toContain('filterColumn');
  });
});

describe('files written by other tools', () => {
  it('survive a Workbook round trip with heights and hidden rows and columns', () => {
    const body =
      '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="10" hidden="true" customWidth="1"/></cols><sheetData><row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1"><c r="A2"><v>2</v></c></row><row r="4" hidden="1" ht="22.5" customHeight="1"><c r="A4"><v>4</v></c></row></sheetData>';
    const bytes = buildXlsx(body);
    const sheet = readFirstSheet(Workbook.fromBuffer(bytes).toBuffer());
    expect(sheet.rowHeights).toEqual({ 0: 30, 3: 22.5 });
    expect(sheet.hiddenRows).toEqual([1, 3]);
    expect(sheet.hiddenColumns).toEqual([1]);
    expect(cellAt(sheet, 'A4')?.value).toBe(4);
  });
});
