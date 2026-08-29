import { XML_NS } from './constants';
import { StyleManager, normalizeColor } from './style-manager';
import {
  dateToExcelSerial,
  isDate,
  validateRowIndex,
  validateColIndex,
  validateCellValue,
} from './date-utils';
import { calculateColumnWidths, generateColsXml } from './column-width';
import { CellValue, CellValidation, CellStyle, ConditionalFormat } from './types';

export type { CellValidation, CellStyle } from './types';

export const indexToColumnLetter = (index: number): string => {
  let letter = '';
  let num = index + 1;

  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }

  return letter;
};

export interface SheetGenerationOptions {
  freezePane?: { row?: number; col?: number };
  autoWidth?: boolean;
  columnWidths?: number[];
  mergeCells?: string[];
  conditionalFormats?: ConditionalFormat[];
  sharedStrings?: Map<string, number>;
}

const CF_OPERATOR_XML: Record<string, string> = {
  greaterThan: 'greaterThan',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
  equal: 'equal',
  notEqual: 'notEqual',
  between: 'between',
  notBetween: 'notBetween',
};

const cfFormulaValue = (value: number | string): string =>
  typeof value === 'number' ? String(value) : `&quot;${escapeXml(value)}&quot;`;

const generateConditionalFormattingXml = (
  formats: ConditionalFormat[] = [],
  styleManager?: StyleManager
): string => {
  if (formats.length === 0) return '';

  return formats
    .map((cf, index) => {
      const priority = index + 1;

      if (cf.type === 'colorScale') {
        const cfvos =
          cf.colors.length === 3
            ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
            : '<cfvo type="min"/><cfvo type="max"/>';
        const colorsXml = cf.colors.map(c => `<color rgb="${normalizeColor(c)}"/>`).join('');
        return `\n  <conditionalFormatting sqref="${cf.range}">\n    <cfRule type="colorScale" priority="${priority}">\n      <colorScale>${cfvos}${colorsXml}</colorScale>\n    </cfRule>\n  </conditionalFormatting>`;
      }

      const dxfId = styleManager ? styleManager.getDxfId(cf.style) : 0;

      if (cf.type === 'expression') {
        return `\n  <conditionalFormatting sqref="${cf.range}">\n    <cfRule type="expression" dxfId="${dxfId}" priority="${priority}">\n      <formula>${escapeXml(cf.formula)}</formula>\n    </cfRule>\n  </conditionalFormatting>`;
      }

      const operator = CF_OPERATOR_XML[cf.operator];
      const formulasXml =
        cf.operator === 'between' || cf.operator === 'notBetween'
          ? `<formula>${cfFormulaValue(cf.value)}</formula><formula>${cfFormulaValue(cf.value2!)}</formula>`
          : `<formula>${cfFormulaValue(cf.value)}</formula>`;
      return `\n  <conditionalFormatting sqref="${cf.range}">\n    <cfRule type="cellIs" dxfId="${dxfId}" priority="${priority}" operator="${operator}">${formulasXml}</cfRule>\n  </conditionalFormatting>`;
    })
    .join('');
};

export const generateRowXml = (
  row: CellValue[],
  rowIndex: number,
  styles: Record<string, CellStyle> = {},
  styleManager?: StyleManager,
  sharedStrings?: Map<string, number>
): string => {
  validateRowIndex(rowIndex);
  let rowXml = `\n    <row r="${rowIndex + 1}">`;

  row.forEach((cellValue, colIndex) => {
    validateColIndex(colIndex);
    const ref = `${indexToColumnLetter(colIndex)}${rowIndex + 1}`;
    const styleKey = `${rowIndex}-${colIndex}`;
    const cellStyle = styles[styleKey];

    let cellXml = `<c r="${ref}"`;

    if (cellStyle && styleManager) {
      const styleId = styleManager.getStyleId(cellStyle);
      cellXml += ` s="${styleId}"`;
    }

    if (cellValue === null || cellValue === undefined) {
      rowXml += cellXml + '/>';
      return;
    }

    if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
      const formula = escapeXml(cellValue.substring(1));
      cellXml += `><f>${formula}</f></c>`;
      rowXml += cellXml;
      return;
    }

    if (isDate(cellValue)) {
      const serial = dateToExcelSerial(cellValue);
      const dateStyleId = styleManager ? styleManager.getDateStyleId() : 0;
      cellXml = `<c r="${ref}" s="${dateStyleId}"><v>${serial}</v></c>`;
      rowXml += cellXml;
      return;
    }

    if (typeof cellValue === 'number') {
      cellXml += `><v>${cellValue}</v></c>`;
      rowXml += cellXml;
      return;
    }

    if (typeof cellValue === 'boolean') {
      cellXml += ` t="b"><v>${cellValue ? 1 : 0}</v></c>`;
      rowXml += cellXml;
      return;
    }

    const stringValue = cellValue.toString();
    validateCellValue(stringValue);

    if (sharedStrings) {
      const index = sharedStrings.get(stringValue);
      if (index !== undefined) {
        cellXml += ` t="s"><v>${index}</v></c>`;
        rowXml += cellXml;
        return;
      }
    }

    const space = stringValue !== stringValue.trim() ? ' xml:space="preserve"' : '';
    cellXml += ` t="inlineStr"><is><t${space}>${escapeXml(stringValue)}</t></is></c>`;
    rowXml += cellXml;
  });

  rowXml += `</row>`;
  return rowXml;
};

export const generateSheetXml = (
  data: CellValue[][],
  validations: CellValidation[] = [],
  styles: Record<string, CellStyle> = {},
  styleManager?: StyleManager,
  options: SheetGenerationOptions = {}
) => {
  let rowsXml = '';

  data.forEach((row, rowIndex) => {
    rowsXml += generateRowXml(row, rowIndex, styles, styleManager, options.sharedStrings);
  });

  let validationsXml = '';
  if (validations.length > 0) {
    validationsXml = `
  <dataValidations count="${validations.length}">`;
    validations.forEach(v => {
      const type = v.type ?? 'list';
      const allowBlank = v.allowBlank === false ? '0' : '1';

      if (type === 'list') {
        const formula1 = v.formula1 ?? `"${escapeXml(v.options)}"`;
        validationsXml += `
    <dataValidation type="list" allowBlank="${allowBlank}" showInputMessage="1" showErrorMessage="1" sqref="${v.range}">
      <formula1>${formula1}</formula1>
    </dataValidation>`;
        return;
      }

      const operator = v.operator ?? 'between';
      let formulas = '';
      if (v.formula1 !== undefined) formulas += `<formula1>${escapeXml(v.formula1)}</formula1>`;
      if (v.formula2 !== undefined) formulas += `<formula2>${escapeXml(v.formula2)}</formula2>`;
      validationsXml += `
    <dataValidation type="${type}" operator="${operator}" allowBlank="${allowBlank}" showInputMessage="1" showErrorMessage="1" sqref="${v.range}">${formulas}</dataValidation>`;
    });
    validationsXml += `
  </dataValidations>`;
  }

  const colsXml = options.columnWidths
    ? generateColsXml(options.columnWidths)
    : options.autoWidth
      ? generateColsXml(calculateColumnWidths(data))
      : '';

  let sheetViewsXml = '';
  if (options.freezePane) {
    const { row = 0, col = 0 } = options.freezePane;
    const topLeftCell = `${indexToColumnLetter(col)}${row + 1}`;
    sheetViewsXml = `  <sheetViews>
    <sheetView workbookViewId="0">`;

    if (row > 0 || col > 0) {
      let activePane = 'bottomRight';
      if (row > 0 && col === 0) {
        activePane = 'bottomLeft';
      } else if (col > 0 && row === 0) {
        activePane = 'topRight';
      }

      sheetViewsXml += `
      <pane`;
      if (col > 0) sheetViewsXml += ` xSplit="${col}"`;
      if (row > 0) sheetViewsXml += ` ySplit="${row}"`;
      sheetViewsXml += ` topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`;
    }

    sheetViewsXml += `
    </sheetView>
  </sheetViews>`;
  }

  let mergeCellsXml = '';
  if (options.mergeCells && options.mergeCells.length > 0) {
    mergeCellsXml = `\n  <mergeCells count="${options.mergeCells.length}">`;
    options.mergeCells.forEach(range => {
      mergeCellsXml += `\n    <mergeCell ref="${range}"/>`;
    });
    mergeCellsXml += `\n  </mergeCells>`;
  }

  const conditionalFormattingXml = generateConditionalFormattingXml(
    options.conditionalFormats,
    styleManager
  );

  return `<?xml version="1.0"?>
<worksheet xmlns="${XML_NS.spreadsheetml}">${sheetViewsXml}${colsXml}
  <sheetData>${rowsXml}
  </sheetData>${validationsXml}${mergeCellsXml}${conditionalFormattingXml}
</worksheet>`;
};

export const generateSharedStringsXml = (strings: string[]) => {
  const uniqueStrings = [...new Set(strings)];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${XML_NS.spreadsheetml}" count="${strings.length}" uniqueCount="${uniqueStrings.length}">
  ${uniqueStrings
    .map(str => {
      const space = str !== str.trim() ? ' xml:space="preserve"' : '';
      return `<si><t${space}>${escapeXml(str)}</t></si>`;
    })
    .join('')}
</sst>`;
};

export const generateStylesXml = (styleManager?: StyleManager) => {
  if (!styleManager) {
    return `<?xml version="1.0"?>
<styleSheet xmlns="${XML_NS.spreadsheetml}">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }

  const numFmtsCount = styleManager.getNumFmtsCount();
  const numFmtsXml =
    numFmtsCount > 0
      ? `  <numFmts count="${numFmtsCount}">\n${styleManager.generateNumFmtsXml()}\n  </numFmts>\n`
      : '';

  return `<?xml version="1.0"?>
<styleSheet xmlns="${XML_NS.spreadsheetml}">
${numFmtsXml}  <fonts count="${styleManager.getFontsCount()}">
${styleManager.generateFontsXml()}
  </fonts>
  <fills count="${styleManager.getFillsCount()}">
${styleManager.generateFillsXml()}
  </fills>
  <borders count="${styleManager.getBordersCount()}">
${styleManager.generateBordersXml()}
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="${styleManager.getCellXfsCount()}">
${styleManager.generateCellXfsXml()}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="${styleManager.getDxfsCount()}">
${styleManager.generateDxfsXml()}
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
};

export const generateContentTypesXml = (
  sheetCount: number = 1,
  hasSharedStrings: boolean = false
) => {
  const sharedStringsOverride = hasSharedStrings
    ? '\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    : '';

  const worksheetOverrides = Array.from({ length: sheetCount }, (_, i) => {
    const sheetNum = i + 1;
    return `  <Override PartName="/xl/worksheets/sheet${sheetNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${XML_NS.content_types}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${worksheetOverrides}${sharedStringsOverride}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
};

export const generateWorkbookXml = (sheetNames: string[] = ['Sheet1']) => {
  const sheetsXml = sheetNames
    .map((name, index) => {
      const sheetId = index + 1;
      const rId = `rId${sheetId}`;
      return `    <sheet name="${escapeXml(name)}" sheetId="${sheetId}" r:id="${rId}"/>`;
    })
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${XML_NS.spreadsheetml}" xmlns:r="${XML_NS.relationships}">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16925"/>
  <workbookPr defaultThemeVersion="166925"/>
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="22260" windowHeight="12645"/>
  </bookViews>
  <calcPr calcId="162913" fullCalcOnLoad="1"/>
  <sheets>
${sheetsXml}
  </sheets>
</workbook>`;
};

export const generateWorkbookRelsXml = (
  sheetCount: number = 1,
  hasSharedStrings: boolean = false
) => {
  const worksheetRels = Array.from({ length: sheetCount }, (_, i) => {
    const rId = i + 1;
    return `  <Relationship Id="rId${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${rId}.xml"/>`;
  }).join('\n');

  const stylesRId = sheetCount + 1;
  const sharedStringsRId = sheetCount + 2;

  return `<?xml version="1.0"?>
<Relationships xmlns="${XML_NS.main_rel}">
${worksheetRels}
  <Relationship Id="rId${stylesRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${hasSharedStrings ? `\n  <Relationship Id="rId${sharedStringsRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` : ''}
</Relationships>`;
};

export const generateCorePropsXml = (
  creator: string = 'Excel Bridge',
  title?: string,
  subject?: string
) => {
  const now = new Date().toISOString();
  const titleXml = title ? `\n  <dc:title>${escapeXml(title)}</dc:title>` : '';
  const subjectXml = subject ? `\n  <dc:subject>${escapeXml(subject)}</dc:subject>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${escapeXml(creator)}</dc:creator>
  <cp:lastModifiedBy>${escapeXml(creator)}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>${titleXml}${subjectXml}
</cp:coreProperties>`;
};

export const generateAppPropsXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Excel Bridge</Application>
  <AppVersion>1.0</AppVersion>
</Properties>`;
};

export const generateRootRelsXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NS.main_rel}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
};

const escapeXml = (text: string): string => {
  return text
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
};
