import { describe, it, expect } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { ExcelBridge, ExcelReader } from '../src';
import { DEFAULT_READER_LIMITS } from '../src/reader';
import { REL_NS, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const corruptEntry = (zip: Uint8Array, entryName: string): Uint8Array => {
  const copy = new Uint8Array(zip);
  const view = new DataView(copy.buffer);
  const nameBytes = strToU8(entryName);
  for (let offset = 0; offset < copy.length - 30; offset++) {
    if (view.getUint32(offset, true) !== 0x04034b50) continue;
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const name = copy.subarray(offset + 30, offset + 30 + nameLength);
    if (name.length !== nameBytes.length || name.some((byte, i) => byte !== nameBytes[i])) continue;
    const dataStart = offset + 30 + nameLength + extraLength;
    copy.fill(0xff, dataStart, dataStart + view.getUint32(offset + 18, true));
    return copy;
  }
  throw new Error(`entry ${entryName} not found`);
};

const singleCellSheet = (ref: string, row: number) =>
  `<sheetData><row r="${row}"><c r="${ref}"><v>1</v></c></row></sheetData>`;

describe('reader grid limits', () => {
  it('rejects a column beyond XFD instead of allocating', () => {
    const xlsx = buildXlsx(singleCellSheet('ZZZZZZZ1', 1));
    expect(() => ExcelBridge.read(xlsx)).toThrow(/Column index .* exceeds Excel limit/);
  });

  it('rejects a column one past XFD', () => {
    const xlsx = buildXlsx(singleCellSheet('XFE1', 1));
    expect(() => ExcelBridge.read(xlsx)).toThrow(/Column index 16384 exceeds Excel limit/);
  });

  it('rejects a row beyond 1048576', () => {
    const xlsx = buildXlsx(singleCellSheet('A1048577', 1048577));
    expect(() => ExcelBridge.read(xlsx)).toThrow(/Row index 1048576 exceeds Excel limit/);
  });

  it('reads the last cell of the grid', () => {
    const xlsx = buildXlsx(singleCellSheet('XFD1048576', 1048576));
    const [sheet] = ExcelBridge.read(xlsx).sheets;
    expect(sheet.data).toHaveLength(1048576);
    expect(Object.keys(sheet.data)).toEqual(['1048575']);
    expect(sheet.data[1048575]).toHaveLength(16384);
    expect(sheet.data[1048575][16383].value).toBe(1);
    expect(sheet.data[1048575][16383].coordinate).toBe('XFD1048576');
  });

  it('clamps a huge <col> range to the grid', () => {
    const xlsx = buildXlsx(
      '<cols><col min="1" max="100000000" width="10"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
    );
    const [sheet] = ExcelBridge.read(xlsx).sheets;
    expect(sheet.columnWidths).toHaveLength(16384);
    expect(sheet.columnWidths?.every(width => width === 10)).toBe(true);
  });

  it('keeps a full-grid <col> range intact', () => {
    const xlsx = buildXlsx(
      '<cols><col min="1" max="16384" width="9"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
    );
    expect(ExcelBridge.read(xlsx).sheets[0].columnWidths).toHaveLength(16384);
  });

  it('ignores a <col> range that starts past the grid', () => {
    const xlsx = buildXlsx(
      '<cols><col min="16385" max="20000" width="9"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
    );
    expect(ExcelBridge.read(xlsx).sheets[0].columnWidths).toHaveLength(0);
  });
});

describe('reader placeholder budget', () => {
  const wideSparseRows = (rows: number) => {
    let body = '<sheetData>';
    for (let r = 1; r <= rows; r++) {
      body += `<row r="${r}"><c r="XFD${r}"><v>1</v></c></row>`;
    }
    return `${body}</sheetData>`;
  };

  it('rejects sheets that pad more empty cells than the budget', () => {
    const rows = Math.ceil(DEFAULT_READER_LIMITS.maxCells / 16383) + 1;
    const xlsx = buildXlsx(wideSparseRows(rows));
    expect(() => ExcelBridge.read(xlsx)).toThrow(/Workbook has at least \d+ cells/);
  });

  it('accepts sheets under the budget', () => {
    const xlsx = buildXlsx(wideSparseRows(50));
    expect(ExcelBridge.read(xlsx).sheets[0].data).toHaveLength(50);
  });
});

describe('reader selective inflation', () => {
  const unusedEntries = {
    'xl/media/image1.png': new Uint8Array(100).fill(7),
    'xl/worksheets/pad.bin': new Uint8Array(100).fill(7),
    'xl/worksheets/junk.xml': new Uint8Array(100).fill(7),
    'xl/worksheets/_rels/junk.xml.rels': new Uint8Array(100).fill(7),
  };

  it('does not inflate entries the workbook does not reference', () => {
    let poisoned = buildXlsx(singleCellSheet('A1', 1), unusedEntries);
    for (const name of Object.keys(unusedEntries)) poisoned = corruptEntry(poisoned, name);
    expect(() => unzipSync(poisoned)).toThrow();

    const [sheet] = new ExcelReader().parseFromBuffer(poisoned).sheets;
    expect(sheet.data[0][0].value).toBe(1);
  });

  it('still fails when a part the reader needs is corrupt', () => {
    const corrupted = corruptEntry(buildXlsx(singleCellSheet('A1', 1)), 'xl/worksheets/sheet1.xml');
    expect(() => ExcelBridge.read(corrupted)).toThrow(/Failed to parse Excel file/);
  });

  const twoSheetWorkbook = (secondTarget: string, secondPath: string) =>
    zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'xl/workbook.xml': strToU8(
        `<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Std" sheetId="1" r:id="rId1"/><sheet name="Odd" sheetId="2" r:id="rId2"/></sheets></workbook>`
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="worksheet" Target="${secondTarget}"/></Relationships>`
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        `<worksheet xmlns="${SPREADSHEET_NS}">${singleCellSheet('A1', 1)}</worksheet>`
      ),
      [secondPath]: strToU8(
        `<worksheet xmlns="${SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>`
      ),
    });

  it('loads sheets whose relationship target is outside xl/worksheets', () => {
    const { sheets } = ExcelBridge.read(twoSheetWorkbook('custom/odd.xml', 'xl/custom/odd.xml'));
    expect(sheets.map(sheet => sheet.name)).toEqual(['Std', 'Odd']);
    expect(sheets[1].data[0][0].value).toBe(2);
  });

  it('loads sheets referenced by an absolute target', () => {
    const { sheets } = ExcelBridge.read(
      twoSheetWorkbook('/xl/worksheets/second.xml', 'xl/worksheets/second.xml')
    );
    expect(sheets.map(sheet => sheet.name)).toEqual(['Std', 'Odd']);
    expect(sheets[1].data[0][0].value).toBe(2);
  });

  it('loads a workbook whose only sheet lives in a worksheets subfolder', () => {
    const xlsx = zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'xl/workbook.xml': strToU8(
        `<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="worksheet" Target="worksheets/nested/sheet1.xml"/></Relationships>`
      ),
      'xl/worksheets/nested/sheet1.xml': strToU8(
        `<worksheet xmlns="${SPREADSHEET_NS}">${singleCellSheet('A1', 1)}</worksheet>`
      ),
    });
    expect(ExcelBridge.read(xlsx).sheets[0].data[0][0].value).toBe(1);
  });

  it('rejects a package without a workbook part', () => {
    const xlsx = zipSync({ '[Content_Types].xml': strToU8('<Types/>') });
    expect(() => ExcelBridge.read(xlsx)).toThrow(/Invalid Excel file structure/);
  });
});
