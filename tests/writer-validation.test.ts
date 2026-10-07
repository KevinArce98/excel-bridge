import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import {
  ExcelBridge,
  ExcelWriter,
  Workbook,
  calculateColumnWidths,
  createExcelWorkbookStream,
  streamToBuffer,
} from '../src';
import type { ConditionalFormat, StreamingSheetInput } from '../src';
import { cellAt, readFirstSheet } from './helpers/read';

type Sheet = Parameters<ExcelWriter['createWorkbookBuffer']>[0][number];

const writer = new ExcelWriter();
const write = (...sheets: Sheet[]) => writer.createWorkbookBuffer(sheets);
const streamOf = (...sheets: StreamingSheetInput[]) =>
  streamToBuffer(createExcelWorkbookStream(sheets));
const named = (name: string): Sheet => ({ data: [['a']], options: { name } });

describe('sheet names', () => {
  it('accepts 31 characters, quotes inside the name and markup characters', () => {
    expect(() => write(named('x'.repeat(31)))).not.toThrow();
    expect(() => write(named("O'Brien"))).not.toThrow();
    expect(() => write(named('Q1 "Sales" & <Co>'))).not.toThrow();
  });

  it('rejects a name longer than 31 characters and quotes it', () => {
    expect(() => write(named('x'.repeat(32)))).toThrow(`Invalid sheet name "${'x'.repeat(32)}"`);
  });

  it.each(['\\', '/', '?', '*', '[', ']', ':'])('rejects a name containing %s', character => {
    expect(() => write(named(`bad${character}name`))).toThrow(/Invalid sheet name/);
  });

  it.each(["'quoted", "quoted'"])('rejects the name %s', name => {
    expect(() => write(named(name))).toThrow(/no apostrophe at either end/);
  });

  it('rejects two sheets with the same name, whatever the case', () => {
    expect(() => write(named('Same'), named('Same'))).toThrow(/used twice/);
    expect(() => write(named('Same'), named('same'))).toThrow(/used twice/);
  });

  it('rejects a name that collides with the default name of another sheet', () => {
    expect(() => write(named('Sheet2'), { data: [['b']] })).toThrow(/"Sheet2" is used twice/);
  });

  it('rejects a workbook without sheets', () => {
    expect(() => write()).toThrow('A workbook needs at least one sheet');
    expect(() => Workbook.create().toBuffer()).toThrow('A workbook needs at least one sheet');
  });

  it('checks the streaming writer the same way', async () => {
    const rows = [['a']];
    await expect(streamOf({ name: 'x'.repeat(40), rows })).rejects.toThrow(/Invalid sheet name/);
    await expect(streamOf({ name: 'a:b', rows })).rejects.toThrow(/Invalid sheet name/);
    await expect(streamOf({ name: 'A', rows }, { name: 'a', rows })).rejects.toThrow(/used twice/);
    await expect(streamOf()).rejects.toThrow('A workbook needs at least one sheet');
  });

  it('checks the names given to Workbook.addSheet', () => {
    const workbook = Workbook.create();
    workbook.addSheet('Data');
    expect(() => workbook.addSheet('data')).toThrow('Sheet "data" already exists');
    expect(() => workbook.addSheet('bad/name')).toThrow(/Invalid sheet name/);
    expect(() => workbook.addSheet('x'.repeat(32))).toThrow(/Invalid sheet name/);
    expect(() => workbook.addSheet('')).toThrow(/Invalid sheet name/);
    expect(workbook.getSheetNames()).toEqual(['Data']);
  });
});

describe('colours', () => {
  const colorOf = (color: string): Sheet => ({
    data: [['a']],
    styles: { '0-0': { color, background: color } },
  });

  it.each(['#FFF', '#FF0000', 'ff0000', '#80FF0000'])('accepts %s', color => {
    expect(() => write(colorOf(color))).not.toThrow();
  });

  it.each(['red', '#GGGGGG', '#12345', '#FF00000000', 'rgb(255,0,0)'])('rejects %j', color => {
    expect(() => write(colorOf(color))).toThrow(/Invalid colour/);
  });

  it('checks the colours of conditional formats and colour scales', () => {
    const rules: ConditionalFormat[] = [
      { type: 'expression', range: 'A1', formula: 'TRUE', style: { background: 'red' } },
      { type: 'colorScale', range: 'A1', colors: ['#FF0000', 'blue'] },
    ];
    for (const rule of rules) {
      expect(() => write({ data: [['a']], conditionalFormats: [rule] })).toThrow(/Invalid colour/);
    }
  });

  it('checks the colours of a streamed sheet', async () => {
    await expect(streamOf({ rows: [['a']], styles: { '0-0': { color: 'red' } } })).rejects.toThrow(
      /Invalid colour/
    );
  });
});

describe('values a worksheet cannot store', () => {
  it.each([NaN, Infinity, -Infinity])('rejects %s and names the cell', value => {
    expect(() =>
      write({
        data: [
          ['a', 'b'],
          ['c', value],
        ],
      })
    ).toThrow(new RegExp(`Cell B2 holds ${value}`));
  });

  it('rejects an invalid Date and names the cell', () => {
    expect(() => write({ data: [[new Date('not a date')]] })).toThrow(
      'Cell A1 holds an invalid Date'
    );
  });

  it('rejects them in a streamed sheet', async () => {
    await expect(streamOf({ rows: [['a'], [NaN]] })).rejects.toThrow('Cell A2 holds NaN');
  });

  it('still writes zero, negative zero, extremes and valid dates', () => {
    const sheet = readFirstSheet(
      write({
        data: [[0, -0, Number.MAX_VALUE, Number.MIN_VALUE, -1.5, new Date(2024, 0, 15)]],
      })
    );
    expect(sheet.data[0].map(cell => cell.type)).toEqual([
      'number',
      'number',
      'number',
      'number',
      'number',
      'date',
    ]);
    expect(cellAt(sheet, 'C1')?.value).toBe(Number.MAX_VALUE);
  });
});

describe('column widths', () => {
  const tall = (rows: number) => Array.from({ length: rows }, () => ['x']);

  it('measures one million rows', () => {
    expect(calculateColumnWidths(tall(1_000_000))).toEqual([8]);
  });

  it('measures a sheet written with autoWidth past the old call stack limit', () => {
    const bytes = write({ data: tall(150_000), options: { autoWidth: true } });
    expect(readFirstSheet(bytes).columnWidths).toEqual([8]);
  });

  it('measures ragged rows, empty data and rows with holes', () => {
    expect(calculateColumnWidths([])).toEqual([]);
    expect(calculateColumnWidths([['a'], ['a', 'b', 'c']])).toEqual([8, 8, 8]);
    expect(calculateColumnWidths([['long text that needs room'], ['a']])).toEqual([33]);
    const sparse: string[][] = [];
    sparse[3] = ['wide cell content here'];
    expect(calculateColumnWidths(sparse)).toEqual([29]);
    expect(calculateColumnWidths([new Array(5)])).toEqual([8, 8, 8, 8, 8]);
  });

  it('lets a loaded workbook with row gaps use autoWidth', () => {
    const workbook = Workbook.create();
    workbook.addSheet('S', [['a']]);
    workbook.setCellValue('S', 9, 1, 'a much longer value than the others');
    workbook.setAutoWidth('S', true);
    const widths = readFirstSheet(workbook.toBuffer()).columnWidths;
    expect(widths).toHaveLength(2);
    expect(widths?.[1]).toBeGreaterThan(8);
  });
});

describe('streaming writer zip entries', () => {
  const methods = async (stream: Promise<Uint8Array>) => {
    const found: number[] = [];
    unzipSync(await stream, {
      filter: file => {
        found.push(file.compression);
        return false;
      },
    });
    return found;
  };

  it('deflates every entry, so readers that cannot unpack stored entries with a data descriptor still work', async () => {
    const found = await methods(
      streamOf({
        name: 'S',
        rows: [['a', 1]],
        hyperlinks: [{ range: 'A1', url: 'https://example.com/' }],
      })
    );
    expect(found.length).toBeGreaterThanOrEqual(8);
    expect(found).toEqual(found.map(() => 8));
  });

  it('reads back with the library reader', async () => {
    const bytes = await streamOf({ name: 'S', rows: [['a', 1]] });
    expect(cellAt(ExcelBridge.read(bytes).sheets[0], 'B1')?.value).toBe(1);
  });

  it('is no longer than the buffered writer plus a small margin', async () => {
    const rows = Array.from({ length: 200 }, (_, index) => [`row ${index}`, index]);
    const streamed = await streamOf({ name: 'S', rows });
    const buffered = write({ data: rows, options: { name: 'S' } });
    expect(streamed.length).toBeLessThan(buffered.length * 1.2);
  });
});
