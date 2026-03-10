import { XML_NS, CELL_TYPES } from './constants';

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

export const generateSheetXml = (
  data: any[][],
  validations: CellValidation[] = [],
  styles: Record<string, CellStyle> = {}
) => {
  let rowsXml = '';
  
  data.forEach((row, rowIndex) => {
    rowsXml += `<row r="${rowIndex + 1}">`;
    row.forEach((cellValue, colIndex) => {
      const ref = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
      const styleKey = `${rowIndex}-${colIndex}`;
      const style = styles[styleKey];
      
      let cellXml = `<c r="${ref}"`;
      
      if (style) {
        cellXml += ` s="1"`;
      }
      
      if (cellValue === null || cellValue === undefined) {
        rowsXml += cellXml + '/>';
      } else if (typeof cellValue === 'number') {
        cellXml += ` t="${CELL_TYPES.NUMBER}"><v>${cellValue}</v></c>`;
        rowsXml += cellXml;
      } else if (typeof cellValue === 'boolean') {
        cellXml += ` t="${CELL_TYPES.BOOLEAN}"><v>${cellValue ? 1 : 0}</v></c>`;
        rowsXml += cellXml;
      } else {
        cellXml += ` t="${CELL_TYPES.INLINE_STRING}"><is><t>${escapeXml(cellValue.toString())}</t></is></c>`;
        rowsXml += cellXml;
      }
    });
    rowsXml += `</row>`;
  });

  const validationsXml = validations.length > 0 ? `
    <dataValidations count="${validations.length}">
      ${validations.map(v => `<dataValidation type="list" sqref="${v.range}"><formula1>"${escapeXml(v.options)}"</formula1></dataValidation>`).join('')}
    </dataValidations>` : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${XML_NS.spreadsheetml}">
  <sheetData>${rowsXml}</sheetData>${validationsXml}
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${XML_NS.spreadsheetml}">
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFE0B0"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>`;
};

export const generateContentTypesXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${XML_NS.content_types}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
};

export const generateWorkbookXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${XML_NS.spreadsheetml}">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="${XML_NS.relationships}"/>
  </sheets>
</workbook>`;
};

export const generateWorkbookRelsXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NS.relationships}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
};

export const generateRootRelsXml = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NS.relationships}">
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
