import { XML_NS } from './constants';
import { StyleManager } from './style-manager';
import {
  dateToExcelSerial,
  isDate,
  validateRowIndex,
  validateColIndex,
  validateCellValue,
} from './date-utils';
import { calculateColumnWidths, generateColsXml } from './column-width';

export interface CellValidation {
  range: string;
  options: string;
}

export interface CellStyle {
  background?: string;
  border?: boolean;
  bold?: boolean;
  color?: string;
}

const indexToColumnLetter = (index: number): string => {
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
  mergeCells?: string[];
}

export const generateSheetXml = (
  data: any[][],
  validations: CellValidation[] = [],
  styles: Record<string, CellStyle> = {},
  styleManager?: StyleManager,
  options: SheetGenerationOptions = {}
) => {
  let rowsXml = '';

  data.forEach((row, rowIndex) => {
    validateRowIndex(rowIndex);
    rowsXml += `\n    <row r="${rowIndex + 1}">`;

    row.forEach((cellValue, colIndex) => {
      validateColIndex(colIndex);
      const ref = `${indexToColumnLetter(colIndex)}${rowIndex + 1}`;
      const styleKey = `${rowIndex}-${colIndex}`;
      const cellStyle = styles[styleKey];

      let cellXml = `<c r="${ref}"`;

      // Apply style if present
      if (cellStyle && styleManager) {
        const styleId = styleManager.getStyleId(cellStyle);
        cellXml += ` s="${styleId}"`;
      }

      // Handle null/undefined
      if (cellValue === null || cellValue === undefined) {
        rowsXml += cellXml + '/>';
        return;
      }

      // Handle formulas (strings starting with =)
      if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
        const formula = escapeXml(cellValue.substring(1));
        cellXml += `><f>${formula}</f><v>0</v></c>`;
        rowsXml += cellXml;
        return;
      }

      // Handle dates
      if (isDate(cellValue)) {
        const serial = dateToExcelSerial(cellValue);
        const dateStyleId = styleManager ? styleManager.getDateStyleId() : 0;
        cellXml = `<c r="${ref}" s="${dateStyleId}"><v>${serial}</v></c>`;
        rowsXml += cellXml;
        return;
      }

      // Handle numbers
      if (typeof cellValue === 'number') {
        cellXml += `><v>${cellValue}</v></c>`;
        rowsXml += cellXml;
        return;
      }

      // Handle booleans
      if (typeof cellValue === 'boolean') {
        cellXml += ` t="b"><v>${cellValue ? 1 : 0}</v></c>`;
        rowsXml += cellXml;
        return;
      }

      // Handle strings
      const stringValue = cellValue.toString();
      validateCellValue(stringValue);
      cellXml += ` t="inlineStr"><is><t>${escapeXml(stringValue)}</t></is></c>`;
      rowsXml += cellXml;
    });

    rowsXml += `</row>`;
  });

  // Build validations XML (MUST come after sheetData)
  let validationsXml = '';
  if (validations.length > 0) {
    validationsXml = `
  <dataValidations count="${validations.length}">`;
    validations.forEach(v => {
      validationsXml += `
    <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${v.range}">
      <formula1>"${escapeXml(v.options)}"</formula1>
    </dataValidation>`;
    });
    validationsXml += `
  </dataValidations>`;
  }

  // Generate column widths if auto-width is enabled
  const colsXml = options.autoWidth ? generateColsXml(calculateColumnWidths(data)) : '';

  // Generate freeze pane if specified
  let sheetViewsXml = '';
  if (options.freezePane) {
    const { row = 0, col = 0 } = options.freezePane;
    const topLeftCell = `${indexToColumnLetter(col)}${row + 1}`;
    sheetViewsXml = `  <sheetViews>
    <sheetView workbookViewId="0">`;

    if (row > 0 || col > 0) {
      sheetViewsXml += `
      <pane`;
      if (col > 0) sheetViewsXml += ` xSplit="${col}"`;
      if (row > 0) sheetViewsXml += ` ySplit="${row}"`;
      sheetViewsXml += ` topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/>`;
    }

    sheetViewsXml += `
    </sheetView>
  </sheetViews>`;
  }

  // Generate merge cells if specified
  let mergeCellsXml = '';
  if (options.mergeCells && options.mergeCells.length > 0) {
    mergeCellsXml = `\n  <mergeCells count="${options.mergeCells.length}">`;
    options.mergeCells.forEach(range => {
      mergeCellsXml += `\n    <mergeCell ref="${range}"/>`;
    });
    mergeCellsXml += `\n  </mergeCells>`;
  }

  return `<?xml version="1.0"?>
<worksheet xmlns="${XML_NS.spreadsheetml}">${sheetViewsXml}${colsXml}
  <sheetData>${rowsXml}
  </sheetData>${validationsXml}${mergeCellsXml}
</worksheet>`;
};

export const generateSharedStringsXml = (strings: string[]) => {
  const uniqueStrings = [...new Set(strings)];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${XML_NS.spreadsheetml}" count="${strings.length}" uniqueCount="${uniqueStrings.length}">
  ${uniqueStrings.map(str => `<si><t>${escapeXml(str)}</t></si>`).join('')}
</sst>`;
};

export const generateStylesXml = (styleManager?: StyleManager) => {
  if (!styleManager) {
    // Return default static styles if no StyleManager provided
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
</styleSheet>`;
  }

  // Generate dynamic styles using StyleManager
  return `<?xml version="1.0"?>
<styleSheet xmlns="${XML_NS.spreadsheetml}">
  <fonts count="${styleManager.getFontsCount()}">
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

export const generateCorePropsXml = (creator: string = 'Excel Bridge') => {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${escapeXml(creator)}</dc:creator>
  <cp:lastModifiedBy>${escapeXml(creator)}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
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
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
};
