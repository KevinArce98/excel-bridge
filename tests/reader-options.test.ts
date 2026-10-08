import { describe, it, expect } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import {
  ExcelBridge,
  ExcelBridgeError,
  ExcelReader,
  ExcelWriter,
  Workbook,
  parseExcel,
} from '../src';
import type { ExcelReaderOptions } from '../src';
import { DEFAULT_READER_LIMITS } from '../src/reader';
import { REL_NS, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';
import { declareUncompressedSize } from './helpers/zip-lies';

const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const grid = (rows: number, cols: number): string => {
  let body = '<sheetData>';
  for (let r = 1; r <= rows; r++) {
    body += `<row r="${r}">`;
    for (let c = 0; c < cols; c++) body += `<c r="${String.fromCharCode(65 + c)}${r}"><v>${r}</v></c>`;
    body += '</row>';
  }
  return `${body}</sheetData>`;
};

const failureOf = (action: () => unknown): ExcelBridgeError => {
  try {
    action();
  } catch (error) {
    return error as ExcelBridgeError;
  }
  throw new Error('did not throw');
};

const severalSheets = (count: number, sheetBody = grid(1, 1)): Uint8Array => {
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8('<Types/>'),
    '_rels/.rels': strToU8('<Relationships/>'),
  };
  let sheets = '';
  let rels = '';
  for (let i = 1; i <= count; i++) {
    sheets += `<sheet name="S${i}" sheetId="${i}" r:id="rId${i}"/>`;
    rels += `<Relationship Id="rId${i}" Type="worksheet" Target="worksheets/sheet${i}.xml"/>`;
    entries[`xl/worksheets/sheet${i}.xml`] = strToU8(
      `<worksheet xmlns="${SPREADSHEET_NS}">${sheetBody}</worksheet>`
    );
  }
  entries['xl/workbook.xml'] = strToU8(
    `<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}"><sheets>${sheets}</sheets></workbook>`
  );
  entries['xl/_rels/workbook.xml.rels'] = strToU8(`<Relationships xmlns="${REL_NS}">${rels}</Relationships>`);
  return zipSync(entries);
};

describe('reader options', () => {
  it('keeps new ExcelReader() working with documented defaults', () => {
    expect(DEFAULT_READER_LIMITS).toEqual({
      maxCells: 5_000_000,
      maxPartBytes: 268_435_456,
      maxTotalBytes: 536_870_912,
      maxSheets: Infinity,
    });
    const [sheet] = new ExcelReader().parseFromBuffer(buildXlsx(grid(2, 2))).sheets;
    expect(sheet.data).toHaveLength(2);
  });

  it.each([0, -1, NaN, 'many' as unknown as number])('rejects maxCells %j as INVALID_INPUT', value => {
    const failure = failureOf(() => new ExcelReader({ maxCells: value }));
    expect(failure.code).toBe('INVALID_INPUT');
    expect(failure.message).toBe('maxCells must be a number above 0, or Infinity');
  });

  it('accepts Infinity and ignores undefined', () => {
    const options: ExcelReaderOptions = { maxCells: Infinity, maxPartBytes: undefined };
    expect(() => new ExcelReader(options)).not.toThrow();
  });
});

describe('maxCells', () => {
  const bytes = buildXlsx(grid(2, 2));

  it('allows exactly the limit and rejects one cell more', () => {
    expect(new ExcelReader({ maxCells: 4 }).parseFromBuffer(bytes).sheets[0].data).toHaveLength(2);
    const failure = failureOf(() => new ExcelReader({ maxCells: 3 }).parseFromBuffer(bytes));
    expect(failure.code).toBe('LIMIT_EXCEEDED');
    expect(failure.limit).toBe('maxCells');
    expect(failure.message).toMatch(/^Failed to parse Excel file: Workbook has at least 4 cells, counting the empty cells that pad rows, over the limit maxCells of 3$/);
  });

  it('counts the empty cells that pad a row', () => {
    const sparse = buildXlsx('<sheetData><row r="1"><c r="D1"><v>1</v></c></row></sheetData>');
    expect(new ExcelReader({ maxCells: 4 }).parseFromBuffer(sparse).sheets[0].data[0]).toHaveLength(4);
    expect(failureOf(() => new ExcelReader({ maxCells: 3 }).parseFromBuffer(sparse)).limit).toBe('maxCells');
  });

  it('shares the budget across the sheets of a workbook', () => {
    const two = severalSheets(2, grid(2, 2));
    expect(new ExcelReader({ maxCells: 8 }).parseFromBuffer(two).sheets).toHaveLength(2);
    expect(failureOf(() => new ExcelReader({ maxCells: 7 }).parseFromBuffer(two)).limit).toBe('maxCells');
  });
});

describe('maxPartBytes and maxTotalBytes', () => {
  it('names the part that is over maxPartBytes', () => {
    const bytes = buildXlsx(grid(50, 5));
    const failure = failureOf(() => new ExcelReader({ maxPartBytes: 1000 }).parseFromBuffer(bytes));
    expect(failure.code).toBe('LIMIT_EXCEEDED');
    expect(failure.limit).toBe('maxPartBytes');
    expect(failure.message).toMatch(/Part "xl\/worksheets\/sheet1\.xml" inflates to \d+ bytes/);
    expect(failure.message).toMatch(/over the limit maxPartBytes of 1000$/);
  });

  it('refuses a header that declares 2 GiB before allocating it', () => {
    const bytes = declareUncompressedSize(buildXlsx(grid(1, 1)), 'xl/worksheets/sheet1.xml', 0x7fffffff);
    const arrayBuffersBefore = process.memoryUsage().arrayBuffers;

    const failure = failureOf(() => new ExcelReader().parseFromBuffer(bytes));

    expect(failure.code).toBe('LIMIT_EXCEEDED');
    expect(failure.limit).toBe('maxPartBytes');
    expect(process.memoryUsage().arrayBuffers - arrayBuffersBefore).toBeLessThan(64 * 1024 * 1024);
  });

  it('never inflates more than a header declares, so a lying header cannot grow the allocation', () => {
    const bytes = declareUncompressedSize(buildXlsx(grid(50, 5)), 'xl/worksheets/sheet1.xml', 200);
    const rowsRead = (() => {
      try {
        return new ExcelReader().parseFromBuffer(bytes).sheets[0].data.length;
      } catch (error) {
        expect((error as ExcelBridgeError).code).toBe('INVALID_FILE');
        return 0;
      }
    })();
    expect(rowsRead).toBeLessThan(50);
  });

  it('adds up the parts of both extraction passes against maxTotalBytes', () => {
    const bytes = severalSheets(3, grid(20, 5));
    const one = new ExcelReader({ maxTotalBytes: Infinity }).parseFromBuffer(bytes);
    expect(one.sheets).toHaveLength(3);

    const failure = failureOf(() => new ExcelReader({ maxTotalBytes: 3000 }).parseFromBuffer(bytes));
    expect(failure.code).toBe('LIMIT_EXCEEDED');
    expect(failure.limit).toBe('maxTotalBytes');
  });

  it('does not count parts the reader never inflates', () => {
    const bytes = buildXlsx(grid(1, 1), { 'xl/media/image1.png': new Uint8Array(5_000_000).fill(7) });
    expect(new ExcelReader({ maxTotalBytes: 100_000, maxPartBytes: 100_000 }).parseFromBuffer(bytes).sheets).toHaveLength(1);
  });
});

describe('maxSheets', () => {
  it('rejects before any sheet part is inflated', () => {
    const bytes = severalSheets(3);
    const failure = failureOf(() => new ExcelReader({ maxSheets: 2 }).parseFromBuffer(bytes));
    expect(failure.code).toBe('LIMIT_EXCEEDED');
    expect(failure.limit).toBe('maxSheets');
    expect(failure.message).toMatch(/Workbook has 3 sheets, over the limit maxSheets of 2$/);
  });

  it('allows exactly the limit', () => {
    expect(new ExcelReader({ maxSheets: 3 }).parseFromBuffer(severalSheets(3)).sheets).toHaveLength(3);
  });
});

describe('where the options are accepted', () => {
  const bytes = new ExcelWriter().createWorkbookBuffer([{ data: [['a', 'b'], [1, 2]] }]);

  it.each<[string, (options: ExcelReaderOptions) => unknown]>([
    ['parseExcel', options => parseExcel(bytes, options)],
    ['ExcelBridge.read', options => ExcelBridge.read(bytes, options)],
    ['Workbook.fromBuffer', options => Workbook.fromBuffer(bytes, options)],
  ])('%s enforces them', (_, read) => {
    expect(() => read({ maxCells: 3 })).toThrow(/maxCells/);
    expect(() => read({})).not.toThrow();
  });

  it.each<[string, (file: File, options: ExcelReaderOptions) => Promise<unknown>]>([
    ['ExcelReader.parseFromFile', (file, options) => new ExcelReader(options).parseFromFile(file)],
    ['ExcelBridge.readFromFile', (file, options) => ExcelBridge.readFromFile(file, options)],
    ['Workbook.fromFile', (file, options) => Workbook.fromFile(file, options)],
  ])('%s enforces them', async (_, read) => {
    const file = new File([bytes], 'a.xlsx');
    await expect(read(file, { maxCells: 3 })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    await expect(read(file, {})).resolves.toBeDefined();
  });
});
