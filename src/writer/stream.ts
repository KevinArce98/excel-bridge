import { Zip, ZipDeflate, ZipPassThrough, strToU8 } from 'fflate';
import { XML_NS } from '../core/constants';
import { StyleManager } from '../core/style-manager';
import {
  generateRowXml,
  generateContentTypesXml,
  generateWorkbookXml,
  generateWorkbookRelsXml,
  generateRootRelsXml,
  generateCorePropsXml,
  generateAppPropsXml,
  generateStylesXml,
  indexToColumnLetter,
} from '../core/xml-templates';
import { generateColsXml } from '../core/column-width';
import { CellValue, CellStyle } from '../core/types';
import { ExcelWriterOptions } from './index';

export interface StreamingSheetInput {
  name?: string;
  rows: Iterable<CellValue[]> | AsyncIterable<CellValue[]>;
  styles?: Record<string, CellStyle>;
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  mergeCells?: string[];
}

const buildSheetHeader = (sheet: StreamingSheetInput): string => {
  let header = `<?xml version="1.0"?>\n<worksheet xmlns="${XML_NS.spreadsheetml}">`;

  if (sheet.freezePane) {
    const { row = 0, col = 0 } = sheet.freezePane;
    if (row > 0 || col > 0) {
      const topLeftCell = `${indexToColumnLetter(col)}${row + 1}`;
      let activePane = 'bottomRight';
      if (row > 0 && col === 0) activePane = 'bottomLeft';
      else if (col > 0 && row === 0) activePane = 'topRight';

      header += `  <sheetViews>\n    <sheetView workbookViewId="0">\n      <pane`;
      if (col > 0) header += ` xSplit="${col}"`;
      if (row > 0) header += ` ySplit="${row}"`;
      header += ` topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>\n    </sheetView>\n  </sheetViews>`;
    }
  }

  if (sheet.columnWidths) {
    header += generateColsXml(sheet.columnWidths);
  }

  header += `\n  <sheetData>`;
  return header;
};

const buildSheetFooter = (sheet: StreamingSheetInput): string => {
  let footer = `\n  </sheetData>`;

  if (sheet.mergeCells && sheet.mergeCells.length > 0) {
    footer += `\n  <mergeCells count="${sheet.mergeCells.length}">`;
    for (const range of sheet.mergeCells) {
      footer += `\n    <mergeCell ref="${range}"/>`;
    }
    footer += `\n  </mergeCells>`;
  }

  footer += `\n</worksheet>`;
  return footer;
};

export async function* createExcelWorkbookStream(
  sheets: StreamingSheetInput[],
  options: ExcelWriterOptions = {}
): AsyncGenerator<Uint8Array, void, unknown> {
  const sheetNames = sheets.map((sheet, index) => sheet.name || `Sheet${index + 1}`);
  const styleManager = new StyleManager();
  const pending: Uint8Array[] = [];

  const zip = new Zip((err, chunk) => {
    if (err) throw err;
    if (chunk) pending.push(chunk);
  });

  function* drain(): Generator<Uint8Array> {
    while (pending.length > 0) {
      yield pending.shift()!;
    }
  }

  const addStaticEntry = (path: string, content: string): void => {
    const entry = new ZipPassThrough(path);
    zip.add(entry);
    entry.push(strToU8(content), true);
  };

  addStaticEntry('[Content_Types].xml', generateContentTypesXml(sheets.length, false));
  addStaticEntry('_rels/.rels', generateRootRelsXml());
  addStaticEntry('xl/_rels/workbook.xml.rels', generateWorkbookRelsXml(sheets.length, false));
  addStaticEntry('xl/workbook.xml', generateWorkbookXml(sheetNames));
  yield* drain();

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const sheet = sheets[sheetIndex];
    const entry = new ZipDeflate(`xl/worksheets/sheet${sheetIndex + 1}.xml`, { level: 6 });
    zip.add(entry);

    entry.push(strToU8(buildSheetHeader(sheet)), false);
    yield* drain();

    let rowIndex = 0;
    for await (const row of sheet.rows) {
      const rowXml = generateRowXml(row, rowIndex, sheet.styles, styleManager);
      entry.push(strToU8(rowXml), false);
      yield* drain();
      rowIndex++;
    }

    entry.push(strToU8(buildSheetFooter(sheet)), true);
    yield* drain();
  }

  addStaticEntry('xl/styles.xml', generateStylesXml(styleManager));
  addStaticEntry(
    'docProps/core.xml',
    generateCorePropsXml(options.creator, options.title, options.subject)
  );
  addStaticEntry('docProps/app.xml', generateAppPropsXml());
  yield* drain();

  zip.end();
  yield* drain();
}

export async function streamToBuffer(
  stream: AsyncGenerator<Uint8Array, void, unknown>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;

  for await (const chunk of stream) {
    chunks.push(chunk);
    total += chunk.length;
  }

  const result = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }

  return result;
}
