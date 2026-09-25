import { Zip, ZipDeflate, ZipPassThrough, strToU8 } from 'fflate';
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
  generateSheetViewsXml,
  generateMergeCellsXml,
  generateAutoFilterXml,
  generateHyperlinksXml,
  generateHyperlinkRelsXml,
  generateWorksheetStart,
  generateWorksheetEnd,
  sheetRelsPath,
  filterDatabaseNames,
} from '../core/xml-templates';
import { generateColsXml } from '../core/column-width';
import { prepareHyperlinks, withHyperlinkStyles } from '../core/hyperlinks';
import { AutoFilter, CellValue, CellStyle, Hyperlink } from '../core/types';
import type { ExcelWriterOptions } from './index';

export interface StreamingSheetInput {
  name?: string;
  rows: Iterable<CellValue[]> | AsyncIterable<CellValue[]>;
  styles?: Record<string, CellStyle>;
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  mergeCells?: string[];
  autoFilter?: AutoFilter;
  hyperlinks?: Hyperlink[];
}

const hasFrozenSplit = (freezePane?: { row?: number; col?: number }): boolean =>
  !!freezePane && ((freezePane.row ?? 0) > 0 || (freezePane.col ?? 0) > 0);

export async function* createExcelWorkbookStream(
  sheets: StreamingSheetInput[],
  options: ExcelWriterOptions = {}
): AsyncGenerator<Uint8Array, void, unknown> {
  const sheetNames = sheets.map((sheet, index) => sheet.name || `Sheet${index + 1}`);
  const sheetLinks = sheets.map(sheet => prepareHyperlinks(sheet.hyperlinks));
  const definedNames = filterDatabaseNames(
    sheetNames,
    sheets.map(sheet => sheet.autoFilter)
  );
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
  addStaticEntry('xl/workbook.xml', generateWorkbookXml(sheetNames, definedNames));
  yield* drain();

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const sheet = sheets[sheetIndex];
    const links = sheetLinks[sheetIndex];
    const styles = withHyperlinkStyles(sheet.styles, links);
    const entry = new ZipDeflate(`xl/worksheets/sheet${sheetIndex + 1}.xml`, { level: 6 });
    zip.add(entry);

    const worksheetStart = generateWorksheetStart(
      hasFrozenSplit(sheet.freezePane) ? generateSheetViewsXml(sheet.freezePane) : '',
      sheet.columnWidths ? generateColsXml(sheet.columnWidths) : '',
      links
    );
    entry.push(strToU8(worksheetStart), false);
    yield* drain();

    let rowIndex = 0;
    for await (const row of sheet.rows) {
      const rowXml = generateRowXml(row, rowIndex, styles, styleManager);
      entry.push(strToU8(rowXml), false);
      yield* drain();
      rowIndex++;
    }

    const worksheetEnd = generateWorksheetEnd(
      generateAutoFilterXml(sheet.autoFilter),
      generateMergeCellsXml(sheet.mergeCells),
      '',
      '',
      generateHyperlinksXml(links)
    );
    entry.push(strToU8(worksheetEnd), true);

    const rels = generateHyperlinkRelsXml(links);
    if (rels) {
      addStaticEntry(sheetRelsPath(sheetIndex + 1), rels);
    }
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
