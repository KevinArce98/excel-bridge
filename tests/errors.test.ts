import { describe, it, expect, vi } from 'vitest';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8, zipSync } from 'fflate';
import {
  ExcelBridgeError,
  ExcelBridgeErrorCode,
  ExcelReader,
  ExcelWriter,
  Workbook,
  coordinateToIndex,
  createExcelWorkbookStream,
  dataValidation,
  hyperlink,
  isExcelBridgeError,
  streamToBuffer,
} from '../src';
import type { CellValue } from '../src';
import { buildXlsx } from './helpers/xlsx';

const write = (data: CellValue[][], options = {}) =>
  new ExcelWriter().createWorkbookBuffer([{ data, options }]);

const codeOf = (action: () => unknown): ExcelBridgeErrorCode => {
  try {
    action();
  } catch (error) {
    if (error instanceof ExcelBridgeError) return error.code;
    throw error;
  }
  throw new Error('did not throw');
};

describe('ExcelBridgeError', () => {
  it('is an Error with a name, a code, a message and an optional cause and limit', () => {
    const cause = new Error('inner');
    const error = new ExcelBridgeError('LIMIT_EXCEEDED', 'too big', { cause, limit: 'maxCells' });

    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('ExcelBridgeError');
    expect(error.code).toBe('LIMIT_EXCEEDED');
    expect(error.message).toBe('too big');
    expect(error.cause).toBe(cause);
    expect(error.limit).toBe('maxCells');
    expect(String(error)).toBe('ExcelBridgeError: too big');
  });

  it('has no cause and no limit unless given', () => {
    const error = new ExcelBridgeError('INVALID_INPUT', 'x');
    expect('cause' in error).toBe(false);
    expect('limit' in error).toBe(false);
  });

  it('is recognised by isExcelBridgeError even when it comes from another copy of the library', () => {
    class OtherCopy extends Error {
      name = 'ExcelBridgeError';
      code = 'INVALID_INPUT';
    }
    expect(isExcelBridgeError(new OtherCopy('x'))).toBe(true);
    expect(isExcelBridgeError(new OtherCopy('x') instanceof ExcelBridgeError)).toBe(false);
    expect(isExcelBridgeError(new Error('x'))).toBe(false);
    expect(isExcelBridgeError('ExcelBridgeError')).toBe(false);
    expect(isExcelBridgeError(null)).toBe(false);
  });
});

describe('codes of the errors the library throws', () => {
  it.each<[string, () => unknown, ExcelBridgeErrorCode, RegExp]>([
    ['a NaN cell', () => write([[NaN]]), 'INVALID_INPUT', /Cell A1 holds NaN/],
    ['an invalid Date', () => write([[new Date('x')]]), 'INVALID_INPUT', /invalid Date/],
    ['an unknown error value', () => write([[{ error: '#X' as '#N/A' }]]), 'INVALID_INPUT', /#X/],
    ['an empty formula', () => write([[{ formula: '' }]]), 'INVALID_INPUT', /needs a formula/],
    ['a text over the cell limit', () => write([['x'.repeat(32768)]]), 'INVALID_INPUT', /exceeds Excel limit/],
    ['a bad sheet name', () => write([['x']], { name: 'a/b' }), 'INVALID_INPUT', /Invalid sheet name/],
    ['no sheets', () => new ExcelWriter().createWorkbookBuffer([]), 'INVALID_INPUT', /at least one sheet/],
    [
      'only hidden sheets',
      () => write([['x']], { state: 'hidden' }),
      'INVALID_INPUT',
      /At least one sheet must be visible/,
    ],
    [
      'a colour',
      () => new ExcelWriter().createWorkbookBuffer([{ data: [['x']], styles: { '0-0': { color: 'red' } } }]),
      'INVALID_INPUT',
      /Invalid colour/,
    ],
    [
      'a border style',
      () =>
        new ExcelWriter().createWorkbookBuffer([
          { data: [['x']], styles: { '0-0': { border: 'wavy' as 'thin' } } },
        ]),
      'INVALID_INPUT',
      /Invalid border style/,
    ],
    [
      'a hyperlink',
      () =>
        new ExcelWriter().createWorkbookBuffer([
          { data: [['x']], hyperlinks: [hyperlink.url('A1', 'ftp://x.dev')] },
        ]),
      'INVALID_INPUT',
      /Unsupported hyperlink url/,
    ],
    ['a row height', () => write([['x']], { rowHeights: { 0: 500 } }), 'INVALID_INPUT', /height 500/],
    ['a coordinate', () => coordinateToIndex('1A'), 'INVALID_INPUT', /Invalid coordinate format/],
    ['a range', () => dataValidation.list('A1', []) && write([['x']], { autoFilter: { range: 'A1:' } }), 'INVALID_INPUT', /Invalid range format/],
    ['a missing sheet', () => Workbook.create().getSheetData('x'), 'INVALID_INPUT', /Sheet "x" not found/],
    [
      'a duplicate sheet',
      () => {
        const workbook = Workbook.create();
        workbook.addSheet('a');
        workbook.addSheet('A');
      },
      'INVALID_INPUT',
      /already exists/,
    ],
    ['a row beyond the grid', () => write([]) && new ExcelWriter().createWorkbookBuffer([{ data: Object.assign([], { 1048576: ['x'] }) }]), 'INVALID_INPUT', /Row index 1048576 exceeds/],
  ])('%s is INVALID_INPUT', (_, action, code, message) => {
    expect(codeOf(action)).toBe(code);
    expect(action).toThrow(message);
  });

  it('reports a missing Blob as UNSUPPORTED', () => {
    vi.stubGlobal('Blob', undefined);
    try {
      expect(codeOf(() => new ExcelWriter().createWorkbook([{ data: [['a']] }]))).toBe('UNSUPPORTED');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it.each<[string, Uint8Array]>([
    ['bytes that are not a zip', new Uint8Array([1, 2, 3, 4])],
    ['an empty buffer', new Uint8Array(0)],
    ['a zip without a workbook', zipSync({ 'a.txt': strToU8('x') })],
    ['a cell past column XFD', buildXlsx('<sheetData><row r="1"><c r="XFE1"><v>1</v></c></row></sheetData>')],
    ['a row past 1048576', buildXlsx('<sheetData><row r="1048577"><c r="A1048577"><v>1</v></c></row></sheetData>')],
  ])('reports %s as INVALID_FILE', (_, bytes) => {
    const failure = (() => {
      try {
        new ExcelReader().parseFromBuffer(bytes);
      } catch (error) {
        return error as ExcelBridgeError;
      }
    })();

    expect(failure).toBeInstanceOf(ExcelBridgeError);
    expect(failure?.code).toBe('INVALID_FILE');
    expect(failure?.message).toMatch(/^Failed to parse Excel file: /);
    expect(failure?.cause).toBeInstanceOf(Error);
  });

  it('keeps the original message after the Failed to parse prefix', () => {
    const bytes = buildXlsx('<sheetData><row r="1"><c r="XFE1"><v>1</v></c></row></sheetData>');
    expect(() => new ExcelReader().parseFromBuffer(bytes)).toThrow(
      'Failed to parse Excel file: Column index 16384 exceeds Excel limit (0-16383)'
    );
  });

  it('wraps an error from the XML parser as INVALID_FILE', () => {
    const doctype = '<!DOCTYPE worksheet [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><worksheet/>';
    const bytes = buildXlsx('', { 'xl/worksheets/sheet1.xml': strToU8(doctype) });
    expect(codeOf(() => new ExcelReader().parseFromBuffer(bytes))).toBe('INVALID_FILE');
  });
});

describe('errors in the streaming writer', () => {
  it('rejects up-front validation with INVALID_INPUT before the first chunk', async () => {
    const stream = createExcelWorkbookStream([{ name: 'a/b', rows: [['x']] }]);
    await expect(stream.next()).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('rejects a bad cell with INVALID_INPUT after earlier chunks were yielded', async () => {
    const stream = createExcelWorkbookStream([{ rows: [['ok'], [NaN]] }]);
    const chunks: Uint8Array[] = [];
    let failure: unknown;
    try {
      for await (const chunk of stream) chunks.push(chunk);
    } catch (error) {
      failure = error;
    }
    expect(failure).toBeInstanceOf(ExcelBridgeError);
    expect((failure as ExcelBridgeError).code).toBe('INVALID_INPUT');
    expect(chunks.length).toBeGreaterThan(0);
  });

  it('lets an error from the caller iterable through unchanged', async () => {
    const mine = new TypeError('database went away');
    async function* rows(): AsyncGenerator<CellValue[]> {
      yield ['a'];
      throw mine;
    }
    await expect(streamToBuffer(createExcelWorkbookStream([{ rows: rows() }]))).rejects.toBe(mine);
  });
});

describe('every throw in src is a typed error', () => {
  const walk = (dir: string): string[] =>
    readdirSync(dir).flatMap(name => {
      const path = join(dir, name);
      return statSync(path).isDirectory() ? walk(path) : path.endsWith('.ts') ? [path] : [];
    });

  it('has no bare throw new Error(', () => {
    const root = new URL('../src', import.meta.url).pathname;
    const offenders = walk(root).filter(
      path => !path.endsWith('core/errors.ts') && /throw new Error\(/.test(readFileSync(path, 'utf8'))
    );
    expect(offenders).toEqual([]);
  });
});
