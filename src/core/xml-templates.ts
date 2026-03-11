import { XML_NS } from './constants';

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

export const generateSheetXml = (
  data: any[][],
  validations: CellValidation[] = [],
  styles: Record<string, CellStyle> = {}
) => {
  let rowsXml = '';

  data.forEach((row, rowIndex) => {
    rowsXml += `\n    <row r="${rowIndex + 1}">`;
    row.forEach((cellValue, colIndex) => {
      const ref = `${indexToColumnLetter(colIndex)}${rowIndex + 1}`;
      const styleKey = `${rowIndex}-${colIndex}`;
      const style = styles[styleKey];

      let cellXml = `<c r="${ref}"`;

      if (style) {
        cellXml += ` s="1"`;
      }

      if (cellValue === null || cellValue === undefined) {
        rowsXml += cellXml + '/>';
      } else if (typeof cellValue === 'number') {
        // No type attribute for numbers - Excel assumes numeric by default
        cellXml += `><v>${cellValue}</v></c>`;
        rowsXml += cellXml;
      } else if (typeof cellValue === 'boolean') {
        cellXml += ` t="b"><v>${cellValue ? 1 : 0}</v></c>`;
        rowsXml += cellXml;
      } else {
        cellXml += ` t="inlineStr"><is><t>${escapeXml(cellValue.toString())}</t></is></c>`;
        rowsXml += cellXml;
      }
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

  return `<?xml version="1.0"?>
<worksheet xmlns="${XML_NS.spreadsheetml}">
  <sheetData>${rowsXml}
  </sheetData>${validationsXml}
</worksheet>`;
};

export const generateSharedStringsXml = (strings: string[]) => {
  const uniqueStrings = [...new Set(strings)];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${XML_NS.spreadsheetml}" count="${strings.length}" uniqueCount="${uniqueStrings.length}">
  ${uniqueStrings.map(str => `<si><t>${escapeXml(str)}</t></si>`).join('')}
</sst>`;
};

export const generateStylesXml = () => {
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
};

export const generateContentTypesXml = (hasSharedStrings: boolean = true) => {
  const sharedStringsOverride = hasSharedStrings
    ? '\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    : '';

  return `<?xml version="1.0"?>
<Types xmlns="${XML_NS.content_types}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>${sharedStringsOverride}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
};

export const generateWorkbookXml = () => {
  return `<?xml version="1.0"?>
<workbook xmlns="${XML_NS.spreadsheetml}" xmlns:r="${XML_NS.relationships}">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
};

export const generateWorkbookRelsXml = (hasSharedStrings: boolean = false) => {
  return `<?xml version="1.0"?>
<Relationships xmlns="${XML_NS.main_rel}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${hasSharedStrings ? '\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ''}
</Relationships>`;
};

export const generateRootRelsXml = () => {
  return `<?xml version="1.0"?>
<Relationships xmlns="${XML_NS.main_rel}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
