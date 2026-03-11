import { describe, it, expect } from 'vitest';
import { ExcelBridge } from '../src';
import { unzipSync } from 'fflate';

describe('ZIP Structure Validation', () => {
  it('should create ZIP with correct file paths (no leading slashes)', () => {
    const data = [
      ['Name', 'Age'],
      ['John', 25],
    ];

    const buffer = ExcelBridge.writeBuffer(data);

    // Unzip and check file paths
    const unzipped = unzipSync(buffer);
    const paths = Object.keys(unzipped);

    // Verify no paths start with '/'
    paths.forEach(path => {
      expect(path.startsWith('/')).toBe(false);
    });

    // Verify expected files exist
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

    // First file should be [Content_Types].xml
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
    
    // Verify sharedStrings.xml does NOT exist
    expect(paths).not.toContain('xl/sharedStrings.xml');
    
    // Verify Content_Types.xml doesn't reference sharedStrings
    const contentTypes = new TextDecoder().decode(unzipped['[Content_Types].xml']);
    expect(contentTypes).not.toContain('sharedStrings');
    
    // Verify workbook.xml.rels doesn't reference sharedStrings
    const workbookRels = new TextDecoder().decode(unzipped['xl/_rels/workbook.xml.rels']);
    expect(workbookRels).not.toContain('sharedStrings');
  });

  it('should create valid binary data (no string conversion)', () => {
    const data = [['A', 1]];
    const buffer = ExcelBridge.writeBuffer(data);

    // Verify it's a Uint8Array
    expect(buffer).toBeInstanceOf(Uint8Array);

    // Verify it starts with ZIP signature (PK\x03\x04)
    expect(buffer[0]).toBe(0x50); // 'P'
    expect(buffer[1]).toBe(0x4B); // 'K'
    expect(buffer[2]).toBe(0x03);
    expect(buffer[3]).toBe(0x04);
  });

  it('should return pure Uint8Array (not Buffer or other wrapper)', () => {
    const data = [['Test', 123]];
    const buffer = ExcelBridge.writeBuffer(data);

    // Verify constructor name is exactly 'Uint8Array'
    expect(buffer.constructor.name).toBe('Uint8Array');

    // Verify it has the expected properties
    expect(buffer).toHaveProperty('byteLength');
    expect(buffer).toHaveProperty('buffer');
    expect(buffer.byteLength).toBeGreaterThan(0);

    // Verify it can be passed to Blob constructor without issues
    expect(() => {
      new Blob([buffer], { type: 'application/octet-stream' });
    }).not.toThrow();
  });
});
