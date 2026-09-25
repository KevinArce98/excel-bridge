import { describe, it, expect } from 'vitest';
import { ExcelBridge, ExcelWriter } from '../src';
import { unzipSync } from 'fflate';

describe('ZIP Structure Validation', () => {
  it('should create ZIP with correct file paths (no leading slashes)', () => {
    const data = [
      ['Name', 'Age'],
      ['John', 25],
    ];

    const buffer = ExcelBridge.writeBuffer(data);

    const unzipped = unzipSync(buffer);
    const paths = Object.keys(unzipped);

    paths.forEach(path => {
      expect(path.startsWith('/')).toBe(false);
    });

    expect(paths).toContain('[Content_Types].xml');
    expect(paths).toContain('_rels/.rels');
    expect(paths).toContain('xl/workbook.xml');
    expect(paths).toContain('xl/_rels/workbook.xml.rels');
    expect(paths).toContain('xl/styles.xml');
    expect(paths).toContain('xl/worksheets/sheet1.xml');
  });

  it('should have [Content_Types].xml as first file in ZIP', () => {
    const data = [['Test', 123]];
    const buffer = ExcelBridge.writeBuffer(data);

    const unzipped = unzipSync(buffer);
    const paths = Object.keys(unzipped);

    expect(paths[0]).toBe('[Content_Types].xml');
  });

  it('should NOT generate sharedStrings.xml (we use inlineStr)', () => {
    const data = [
      ['Name', 'Age'],
      ['John', 25],
      ['Jane', 30],
    ];

    const buffer = ExcelBridge.writeBuffer(data);
    const unzipped = unzipSync(buffer);
    const paths = Object.keys(unzipped);

    expect(paths).not.toContain('xl/sharedStrings.xml');

    const contentTypes = new TextDecoder().decode(unzipped['[Content_Types].xml']);
    expect(contentTypes).not.toContain('sharedStrings');

    const workbookRels = new TextDecoder().decode(unzipped['xl/_rels/workbook.xml.rels']);
    expect(workbookRels).not.toContain('sharedStrings');
  });

  it('should create valid binary data (no string conversion)', () => {
    const data = [['A', 1]];
    const buffer = ExcelBridge.writeBuffer(data);

    expect(buffer).toBeInstanceOf(Uint8Array);

    expect(buffer[0]).toBe(0x50);
    expect(buffer[1]).toBe(0x4b);
    expect(buffer[2]).toBe(0x03);
    expect(buffer[3]).toBe(0x04);
  });

  it('should return pure Uint8Array (not Buffer or other wrapper)', () => {
    const data = [['Test', 123]];
    const buffer = ExcelBridge.writeBuffer(data);

    expect(buffer.constructor.name).toBe('Uint8Array');

    expect(buffer).toHaveProperty('byteLength');
    expect(buffer).toHaveProperty('buffer');
    expect(buffer.byteLength).toBeGreaterThan(0);

    expect(() => {
      new Blob([buffer], { type: 'application/octet-stream' });
    }).not.toThrow();
  });

  it('adds sheet relationships only for sheets with external links', () => {
    const buffer = new ExcelWriter().createWorkbookBuffer([
      { data: [['a']], hyperlinks: [{ range: 'A1', url: 'https://example.com' }] },
      { data: [['b']], hyperlinks: [{ range: 'A1', location: "'Sheet1'!A1" }] },
    ]);
    const paths = Object.keys(unzipSync(buffer));

    expect(paths).toContain('xl/worksheets/_rels/sheet1.xml.rels');
    expect(paths).not.toContain('xl/worksheets/_rels/sheet2.xml.rels');
    expect(paths.some(path => path.startsWith('/'))).toBe(false);
    expect(paths[0]).toBe('[Content_Types].xml');
    expect(paths.indexOf('xl/worksheets/_rels/sheet1.xml.rels')).toBe(
      paths.indexOf('xl/worksheets/sheet1.xml') + 1
    );
  });
});
