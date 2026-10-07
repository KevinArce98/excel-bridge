import { strToU8, zipSync } from 'fflate';

export const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const DATE_STYLES = `<?xml version="1.0"?><styleSheet xmlns="${SPREADSHEET_NS}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`;

export interface BuildOptions {
  workbookProperties?: string;
  styles?: string;
}

export const buildXlsx = (
  sheetBody: string,
  extraEntries: Record<string, Uint8Array> = {},
  options: BuildOptions = {}
): Uint8Array =>
  zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    ),
    '_rels/.rels': strToU8(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"/>`),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0"?><workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}">${options.workbookProperties ?? ''}<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<?xml version="1.0"?><worksheet xmlns="${SPREADSHEET_NS}">${sheetBody}</worksheet>`
    ),
    ...(options.styles ? { 'xl/styles.xml': strToU8(options.styles) } : {}),
    ...extraEntries,
  });
