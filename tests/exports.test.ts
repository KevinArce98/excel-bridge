import { describe, expect, it } from 'vitest';
import * as library from '../src';
import {
  ExcelBridge,
  createExcelFile,
  createExcelFileBuffer,
  isExcelErrorValue,
  parseExcel,
} from '../src';

const RUNTIME_EXPORTS = [
  'EXCEL_LIMITS',
  'ExcelBridge',
  'ExcelBridgeError',
  'ExcelReader',
  'ExcelWriter',
  'Workbook',
  'XLSX_CONTENT_TYPE',
  'calculateColumnWidths',
  'coordinateToIndex',
  'createExcelFile',
  'createExcelFileBuffer',
  'createExcelWorkbookStream',
  'dataValidation',
  'dateToExcelSerial',
  'downloadXlsx',
  'excelSerialToDate',
  'hyperlink',
  'indexToCoordinate',
  'isDate',
  'isExcelBridgeError',
  'isExcelErrorValue',
  'objectsToSheet',
  'objectsToStreamingSheet',
  'parseExcel',
  'sheetToObjects',
  'streamToBuffer',
  'toReadableStream',
  'xlsxResponse',
];

const REMOVED_EXPORTS = [
  'StyleManager',
  'generateSheetXml',
  'generateSharedStringsXml',
  'generateStylesXml',
  'generateContentTypesXml',
  'generateWorkbookXml',
  'generateWorkbookRelsXml',
  'generateRootRelsXml',
  'generateCorePropsXml',
  'generateAppPropsXml',
  'generateSheetRelsXml',
  'generateColsXml',
  'createExcelBlob',
  'createExcelBuffer',
  'extractExcelFiles',
  'validateExcelStructure',
  'XML_NS',
  'CONTENT_TYPES',
  'RELATIONSHIP_TYPES',
  'CELL_TYPES',
  'isDateNumFmtId',
  'isDateFormatCode',
  'validateRowIndex',
  'validateColIndex',
  'validateCellValue',
];

describe('public runtime surface', () => {
  it('exports exactly the documented names', () => {
    expect(Object.keys(library).sort()).toEqual(RUNTIME_EXPORTS);
  });

  it.each(REMOVED_EXPORTS)('no longer exports %s', name => {
    expect(library).not.toHaveProperty(name);
  });

  it('no longer has the writer methods that mutated the sheets they were given', () => {
    const writer = new library.ExcelWriter();
    ['addValidation', 'addStyle'].forEach(name => expect(writer).not.toHaveProperty(name));
    ['createSimple', 'createSimpleBuffer'].forEach(name =>
      expect(library.ExcelWriter).not.toHaveProperty(name)
    );
  });
});

describe('shortcuts that stay', () => {
  it('writes with createExcelFileBuffer and reads with parseExcel', () => {
    const sheet = parseExcel(createExcelFileBuffer([['a', 1, { formula: 'B1*2', result: 2 }]]))
      .sheets[0];
    expect(sheet.data[0].map(cell => cell.value)).toEqual(['a', 1, 2]);
    expect(sheet.data[0][2].formula).toBe('B1*2');
  });

  it('writes a Blob with createExcelFile', async () => {
    const blob = createExcelFile([['x']]);
    const sheet = parseExcel(new Uint8Array(await blob.arrayBuffer())).sheets[0];
    expect(sheet.data[0][0].value).toBe('x');
  });

  it('keeps the same shortcuts on ExcelBridge', () => {
    expect(ExcelBridge.read).toBe(parseExcel);
    expect(ExcelBridge.write).toBe(createExcelFile);
    expect(ExcelBridge.writeBuffer).toBe(createExcelFileBuffer);
  });
});

describe('isExcelErrorValue', () => {
  it.each(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'])(
    'accepts %s',
    code => {
      expect(isExcelErrorValue(code)).toBe(true);
    }
  );

  it.each(['#SPILL!', '#n/a', 'N/A', '', null, undefined, 4, { error: '#N/A' }])(
    'rejects %j',
    value => {
      expect(isExcelErrorValue(value)).toBe(false);
    }
  );

  it('takes the value of an error cell to an ErrorCell without a cast', () => {
    const cell = parseExcel(createExcelFileBuffer([[{ error: '#REF!' }]])).sheets[0].data[0][0];
    if (cell.type !== 'error') throw new Error('expected an error cell');
    expect(isExcelErrorValue(cell.value) ? { error: cell.value } : undefined).toEqual({
      error: '#REF!',
    });
  });
});
