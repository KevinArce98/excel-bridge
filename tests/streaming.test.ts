import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { ExcelWriter, createExcelWorkbookStream, streamToBuffer, hyperlink } from '../src';
import type { CellStyle, CellValue, Hyperlink } from '../src';

const data: CellValue[][] = [
  ['Name', 'Site', 'Mail', 'Jump'],
  ['Ann', 'Docs', 'Write', 'Top'],
  ['Bob', 'Blog', 'Ping', 'Top'],
];

const styles: Record<string, CellStyle> = { '0-0': { bold: true }, '1-1': { italic: true } };

const links: Hyperlink[] = [
  { range: 'B2', url: 'https://example.com/docs#intro', tooltip: 'Docs' },
  { range: 'C2', url: 'mailto:ann@example.com' },
  hyperlink.internal('D2', 'Report', 'A1'),
  { range: 'B3', url: 'https://example.com/blog' },
];

describe('Streaming writer', () => {
  it('writes the same parts as ExcelWriter', async () => {
    const writerFiles = unzipSync(
      new ExcelWriter().createWorkbookBuffer([
        {
          data,
          styles,
          hyperlinks: links,
          mergeCells: ['C3:D3'],
          options: {
            name: 'Report',
            freezePane: { row: 1 },
            columnWidths: [10, 20, 20, 10],
            autoFilter: { range: 'A1:D3' },
          },
        },
      ])
    );
    const streamFiles = unzipSync(
      await streamToBuffer(
        createExcelWorkbookStream([
          {
            name: 'Report',
            rows: data,
            styles,
            hyperlinks: links,
            mergeCells: ['C3:D3'],
            freezePane: { row: 1 },
            columnWidths: [10, 20, 20, 10],
            autoFilter: { range: 'A1:D3' },
          },
        ])
      )
    );

    for (const path of [
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      'xl/workbook.xml',
      'xl/styles.xml',
      '[Content_Types].xml',
    ]) {
      expect(strFromU8(streamFiles[path]), path).toBe(strFromU8(writerFiles[path]));
    }
  });

  it('rejects invalid input before yielding anything', async () => {
    const stream = createExcelWorkbookStream([
      { rows: [[1]], hyperlinks: [{ range: 'A1', url: 'javascript:void(0)' }] },
    ]);

    await expect(stream.next()).rejects.toThrow('Unsupported hyperlink url at A1');
  });

  it('rejects an invalid filter range before yielding anything', async () => {
    const stream = createExcelWorkbookStream([{ rows: [[1]], autoFilter: { range: 'A1:A0' } }]);

    await expect(stream.next()).rejects.toThrow('Invalid range format: A1:A0');
  });

  it('writes each sheet relationship part right after its sheet', async () => {
    const buffer = await streamToBuffer(
      createExcelWorkbookStream([
        { rows: data, hyperlinks: links },
        { rows: data, hyperlinks: [{ range: 'A1', url: 'https://example.com' }] },
      ])
    );
    const paths = Object.keys(unzipSync(buffer));

    expect(paths.indexOf('xl/worksheets/_rels/sheet1.xml.rels')).toBe(
      paths.indexOf('xl/worksheets/sheet1.xml') + 1
    );
    expect(paths.indexOf('xl/worksheets/_rels/sheet2.xml.rels')).toBe(
      paths.indexOf('xl/worksheets/sheet2.xml') + 1
    );
  });

  it('keeps [Content_Types].xml as the first entry', async () => {
    const buffer = await streamToBuffer(
      createExcelWorkbookStream([{ rows: data, hyperlinks: links, autoFilter: { range: 'A1:D3' } }])
    );

    expect(Object.keys(unzipSync(buffer))[0]).toBe('[Content_Types].xml');
  });
});
