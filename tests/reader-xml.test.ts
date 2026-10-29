import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { ExcelBridge, ExcelWriter, dataValidation, hyperlink, isExcelBridgeError } from '../src';
import type { ExcelBridgeError } from '../src';
import { cellAt, fixture, readFirstSheet } from './helpers/read';
import { REL_NS, SPREADSHEET_NS, buildXlsx } from './helpers/xlsx';

const mapParts = (bytes: Uint8Array, change: (path: string, xml: string) => string): Uint8Array => {
  const parts = unzipSync(bytes);
  const out: Record<string, Uint8Array> = {};
  for (const [path, content] of Object.entries(parts)) {
    out[path] = /\.(xml|rels)$/.test(path) ? strToU8(change(path, strFromU8(content))) : content;
  }
  return zipSync(out);
};

const prefixElements = (prefix: string) => (_path: string, xml: string) => {
  const renamed = xml.replace(/<(\/?)([A-Za-z]\w*)(?=[\s/>])/g, `<$1${prefix}:$2`);
  return renamed.replace(/<(\w+:\w+)/, `<$1 xmlns:${prefix}="urn:test"`);
};

const richWorkbook = () =>
  new ExcelWriter({ sharedStrings: true, title: 'T & "q"', subject: 'S' }).createWorkbookBuffer([
    {
      data: [
        ['Name', 'Amount', 'Date', 'Link'],
        ['A & B', 12.5, new Date(2024, 1, 29), 'site'],
        ['<tag> "quoted" \'apos\'', 7, null, { formula: 'B2*2', result: 25 }],
        ['line1\nline2', true, { text: '=== not a formula ===' }, { error: '#N/A' }],
      ],
      styles: {
        '0-0': { bold: true, background: '#FFFF00', border: 'thin' },
        '1-1': { numberFormat: '0.00', color: '#FF0000', align: 'right' },
        '1-2': { numberFormat: 'yyyy-mm-dd' },
      },
      mergeCells: ['A5:B6'],
      validations: [
        dataValidation.list('A2:A4', ['x', 'y&z']),
        dataValidation.wholeNumber('B2:B4', 'between', 1, 9),
      ],
      conditionalFormats: [
        {
          type: 'cellValue',
          range: 'B2:B4',
          operator: 'greaterThan',
          value: 5,
          style: { background: '#00FF00' },
        },
        { type: 'colorScale', range: 'B2:B4', colors: ['#FFFFFF', '#FF0000'] },
      ],
      hyperlinks: [hyperlink.url('D2', 'https://example.com/?a=1&b=2', { tooltip: 'go & see' })],
      options: {
        name: 'Data & more',
        freezePane: { row: 1, col: 1 },
        columnWidths: [20, 12],
        autoFilter: { range: 'A1:D4' },
        rowHeights: { 1: 30 },
        hiddenColumns: [3],
      },
    },
    { data: [['second']], options: { name: 'Second', state: 'hidden' } },
  ]);

const source = richWorkbook();

describe('prefixed SpreadsheetML', () => {
  const expected = ExcelBridge.read(source);

  it.each(['x', 'ns0', 'main'])(
    'reads every part with its elements prefixed with %s as the unprefixed file',
    prefix => {
      const bytes = mapParts(source, prefixElements(prefix));
      expect(ExcelBridge.read(bytes)).toEqual(expected);
    }
  );

  it('reads a workbook that only prefixes workbook.xml and its relationships', () => {
    const bytes = mapParts(source, (path, xml) =>
      path === 'xl/workbook.xml' || path === 'xl/_rels/workbook.xml.rels'
        ? prefixElements('x')(path, xml)
        : xml
    );
    expect(ExcelBridge.read(bytes)).toEqual(expected);
  });

  it('reads the hyperlink targets of a prefixed sheet relationships part', () => {
    const bytes = mapParts(source, (path, xml) =>
      path.endsWith('.rels') ? prefixElements('ns0')(path, xml) : xml
    );
    expect(ExcelBridge.read(bytes).sheets[0].hyperlinks).toEqual(expected.sheets[0].hyperlinks);
  });

  it('reads the metadata of a prefixed core.xml', () => {
    expect(expected.metadata).toMatchObject({ title: 'T & "q"', subject: 'S' });
    const bytes = mapParts(source, (path, xml) =>
      path === 'docProps/core.xml'
        ? xml.replace(/cp:/g, 'core:').replace(/xmlns:cp=/, 'xmlns:core=')
        : xml
    );
    expect(ExcelBridge.read(bytes).metadata).toMatchObject({ title: 'T & "q"', subject: 'S' });
  });
});

describe('prefixed SpreadsheetML in a minimal workbook', () => {
  const prefixed = (prefix: string) => {
    const declaration = prefix
      ? `xmlns:${prefix}="${SPREADSHEET_NS}"`
      : `xmlns="${SPREADSHEET_NS}"`;
    const tag = (name: string) => (prefix ? `${prefix}:${name}` : name);
    return zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'xl/workbook.xml': strToU8(
        `<${tag('workbook')} ${declaration} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><${tag('sheets')}><${tag('sheet')} name="S" sheetId="1" r:id="rId1"/></${tag('sheets')}></${tag('workbook')}>`
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        `<${tag('worksheet')} ${declaration}><${tag('sheetData')}><${tag('row')} r="1"><${tag('c')} r="A1"><${tag('v')}>5</${tag('v')}></${tag('c')}></${tag('row')}></${tag('sheetData')}></${tag('worksheet')}>`
      ),
    });
  };

  it.each(['', 'x', 'ns0'])('reads the sheet of a workbook prefixed with "%s"', prefix => {
    expect(cellAt(readFirstSheet(prefixed(prefix)), 'A1')?.value).toBe(5);
  });
});

describe('XML syntax in parts', () => {
  const sheet = (cells: string) => `<sheetData><row r="1">${cells}</row></sheetData>`;
  const text = (xml: string) => cellAt(readFirstSheet(buildXlsx(xml)), 'A1')?.value;

  it('decodes numeric character references in inline strings', () => {
    expect(
      text(sheet('<c r="A1" t="inlineStr"><is><t>caf&#233; &#x1F600; a&#10;b</t></is></c>'))
    ).toBe('café 😀 a\nb');
  });

  it('decodes numeric character references in shared strings, attributes and formulas', () => {
    const sst = `<?xml version="1.0"?><sst xmlns="${SPREADSHEET_NS}"><si><t>&#xD6;ffnen</t></si></sst>`;
    const body = `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>A1&amp;&quot;&#38;&quot;</f></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="S!A1" tooltip="a&#10;b"/></hyperlinks>`;
    const sheetModel = readFirstSheet(buildXlsx(body, { 'xl/sharedStrings.xml': strToU8(sst) }));
    expect(cellAt(sheetModel, 'A1')?.value).toBe('Öffnen');
    expect(cellAt(sheetModel, 'B1')?.formula).toBe('A1&"&"');
    expect(sheetModel.hyperlinks?.[0].tooltip).toBe('a\nb');
  });

  it('reads CDATA and comments inside a string', () => {
    expect(
      text(sheet('<c r="A1" t="inlineStr"><is><t>a<!-- x --><![CDATA[<b>&]]>c</t></is></c>'))
    ).toBe('a<b>&c');
  });

  it('keeps whitespace-only strings that are marked as preserved', () => {
    expect(text(sheet('<c r="A1" t="inlineStr"><is><t xml:space="preserve">  </t></is></c>'))).toBe(
      '  '
    );
  });

  it('reads rich text runs without the phonetic runs', () => {
    const sst = `<sst xmlns="${SPREADSHEET_NS}"><si><r><rPr><b/></rPr><t xml:space="preserve">Hello </t></r><r><t>world</t></r><rPh sb="0" eb="1"><t>ignored</t></rPh><phoneticPr fontId="1"/></si></sst>`;
    const model = readFirstSheet(
      buildXlsx(sheet('<c r="A1" t="s"><v>0</v></c>'), { 'xl/sharedStrings.xml': strToU8(sst) })
    );
    expect(cellAt(model, 'A1')?.value).toBe('Hello world');
  });

  it('ignores mc:AlternateContent and extension lists', () => {
    const body = `<sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting><mc:AlternateContent xmlns:mc="m"><mc:Choice Requires="x14"><x14:conditionalFormattings xmlns:x14="x"><x14:conditionalFormatting><x14:cfRule type="dataBar"/></x14:conditionalFormatting></x14:conditionalFormattings></mc:Choice></mc:AlternateContent><extLst><ext uri="{1}" xmlns:x14="x"><x14:id>{2}</x14:id></ext></extLst>`;
    const model = readFirstSheet(buildXlsx(body));
    expect(cellAt(model, 'A1')?.value).toBe(1);
    expect(model.conditionalFormats).toHaveLength(1);
  });

  it('reads a pretty-printed part as the compact one', () => {
    const bytes = mapParts(source, (_path, xml) => xml.replace(/(<\/[\w:]+>|\/>)(?=<)/g, '$1\n  '));
    expect(ExcelBridge.read(bytes)).toEqual(ExcelBridge.read(source));
  });

  it('reads parts with a BOM, single quotes and CRLF line ends', () => {
    const bytes = mapParts(source, (_path, xml) =>
      ('﻿' + xml)
        .replace(/^﻿<\?xml[^>]*\?>/, `﻿<?xml version='1.0' encoding='UTF-8' standalone='yes'?>`)
        .replace(/\n/g, '\r\n')
    );
    expect(cellAt(ExcelBridge.read(bytes).sheets[0], 'A4')?.value).toBe('line1\nline2');
  });

  it('reads the reference files as before', () => {
    expect(cellAt(ExcelBridge.read(fixture('exceljs-text.xlsx')).sheets[0], 'A8')?.value).toBe(
      '&#233; &amp;'
    );
  });
});

describe('formula string results with preserved whitespace', () => {
  const read = (cell: string) =>
    cellAt(readFirstSheet(buildXlsx(`<sheetData><row r="1">${cell}</row></sheetData>`)), 'A1');

  it('reads the text of a t="str" value that carries xml:space', () => {
    expect(read('<c r="A1" t="str"><v xml:space="preserve"> padded </v></c>')).toMatchObject({
      type: 'string',
      value: ' padded ',
    });
  });

  it('reads a t="str" value of only whitespace', () => {
    expect(read('<c r="A1" t="str"><v xml:space="preserve">  </v></c>')).toMatchObject({
      type: 'string',
      value: '  ',
    });
  });

  it('reads a t="str" value with a formula and a preserved result', () => {
    expect(
      read('<c r="A1" t="str"><f>"a"&amp;" "</f><v xml:space="preserve">a </v></c>')
    ).toMatchObject({ type: 'string', value: 'a ', formula: '"a"&" "' });
  });

  it('reads the padded strings of a SheetJS file', () => {
    const sheet = readFirstSheet(fixture('sheetjs-text.xlsx'));
    expect(
      ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7'].map(coordinate => cellAt(sheet, coordinate)?.value)
    ).toEqual([
      'plain',
      '  padded on both sides  ',
      'trailing ',
      '\ttabbed',
      'a & b < c',
      'line one\nline two',
      ' leading',
    ]);
    expect(cellAt(sheet, 'A7')?.formula).toBe('" leading"');
  });

  it('reads a t="str" value without attributes as before', () => {
    expect(read('<c r="A1" t="str"><v>plain</v></c>')).toMatchObject({
      type: 'string',
      value: 'plain',
    });
  });
});

describe('malformed parts', () => {
  const broken = (path: string, change: (xml: string) => string) =>
    mapParts(source, (current, xml) => (current === path ? change(xml) : xml));

  it.each([
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/sharedStrings.xml',
    'xl/styles.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/_rels/sheet1.xml.rels',
  ])('names %s when it is truncated', path => {
    const bytes = broken(path, xml => xml.slice(0, Math.floor(xml.length * 0.6)));
    expect(() => ExcelBridge.read(bytes)).toThrow(
      new RegExp(`Failed to parse Excel file: ${path.replace(/[.]/g, '\\.')}: Malformed XML`)
    );
  });

  it('throws an INVALID_FILE error that keeps the parser error as its cause', () => {
    const bytes = broken('xl/styles.xml', xml => xml.slice(0, Math.floor(xml.length * 0.6)));
    let thrown: unknown;
    try {
      ExcelBridge.read(bytes);
    } catch (error) {
      thrown = error;
    }
    expect(isExcelBridgeError(thrown)).toBe(true);
    expect(thrown).toMatchObject({ code: 'INVALID_FILE' });
    const cause = (thrown as Error).cause as Error;
    expect(isExcelBridgeError(cause)).toBe(true);
    expect(cause).toMatchObject({ code: 'INVALID_FILE' });
    expect(cause.cause).toBeInstanceOf(Error);
    expect(cause.message).toMatch(/^xl\/styles\.xml: Malformed XML: .* at position \d+$/);
  });

  it('wraps a malformed part once with the path and once with the prefix, keeping the code', () => {
    const bytes = broken('xl/styles.xml', xml => xml.slice(0, Math.floor(xml.length * 0.6)));
    let thrown: unknown;
    try {
      ExcelBridge.read(bytes);
    } catch (error) {
      thrown = error;
    }
    const outer = thrown as ExcelBridgeError;
    const part = outer.cause as ExcelBridgeError;
    const parser = part.cause as ExcelBridgeError;
    expect(outer.message).toMatch(/^Failed to parse Excel file: xl\/styles\.xml: Malformed XML: /);
    expect(outer.message.match(/Failed to parse Excel file/g)).toHaveLength(1);
    expect(part.message).toBe(`xl/styles.xml: ${parser.message}`);
    expect(parser.message).toMatch(/^Malformed XML: /);
    expect([outer.code, part.code, parser.code]).toEqual(['INVALID_FILE', 'INVALID_FILE', 'INVALID_FILE']);
    expect(parser.cause).toBeUndefined();
  });

  it('rejects a part with a DOCTYPE', () => {
    const bytes = broken('xl/sharedStrings.xml', xml =>
      xml.replace('<sst', '<!DOCTYPE sst [<!ENTITY e "boom">]><sst')
    );
    expect(() => ExcelBridge.read(bytes)).toThrow(/xl\/sharedStrings\.xml: Malformed XML: DOCTYPE/);
  });

  it('rejects a mismatched closing tag instead of dropping the cell', () => {
    const bytes = broken('xl/worksheets/sheet1.xml', xml => xml.replace('</c>', '</v>'));
    expect(() => ExcelBridge.read(bytes)).toThrow(/Malformed XML: unexpected closing tag/);
  });

  it('reads the workbook when docProps are malformed and drops their metadata', () => {
    const bytes = broken('docProps/core.xml', xml => xml.slice(0, 80));
    const read = ExcelBridge.read(bytes);
    expect(read.sheets).toHaveLength(2);
    expect(read.metadata.title).toBeUndefined();
  });

  it('keeps the metadata of core.xml when app.xml is malformed', () => {
    const bytes = broken('docProps/app.xml', xml => xml.slice(0, 80));
    expect(ExcelBridge.read(bytes).metadata.title).toBe('T & "q"');
  });
});

describe('parts shaped like Excel output', () => {
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const excelLike = (): Record<string, Uint8Array> => ({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="xl/workbook.xml"/></Relationships>'
    ),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="${MAIN}" xmlns:r="${RELS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/><workbookPr defaultThemeVersion="166925"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="/Users/someone/Documents/" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="8_{AB}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}" xmlns:xr10="a" xmlns:xr6="b"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="17540" xr2:uid="{1}" xmlns:xr2="c"/></bookViews><sheets><sheet name="Q1 &amp; Q2" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="x" Target="sharedStrings.xml"/><Relationship Id="rId2" Type="x" Target="styles.xml"/><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/sharedStrings.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="${MAIN}" count="3" uniqueCount="3"><si><t>plain</t></si><si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">bold </t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>plain</t></r></si><si><t>日本語</t><rPh sb="0" eb="3"><t>ニホンゴ</t></rPh><phoneticPr fontId="1" type="noConversion"/></si></sst>`
    ),
    'xl/styles.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="${MAIN}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><fonts count="1" x14ac:knownFonts="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${MAIN}" xmlns:r="${RELS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{00000000-0001-0000-0000-000000000000}"><dimension ref="A1:C2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C2" sqref="C2"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="16" x14ac:dyDescent="0.2"/><cols><col min="1" max="1" width="23.5" customWidth="1"/></cols><sheetData><row r="1" spans="1:3" x14ac:dyDescent="0.2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2" spans="1:3" x14ac:dyDescent="0.2"><c r="A2" s="1"><v>1234.5</v></c><c r="B2"><f>A2*2</f><v>2469</v></c><c r="C2" t="str"><f>"x"&amp;"y"</f><v>xy</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>`
    ),
  });

  it('reads a workbook with the markup Excel writes', () => {
    const read = ExcelBridge.read(zipSync(excelLike()));
    const [sheet] = read.sheets;
    expect(sheet.name).toBe('Q1 & Q2');
    expect(cellAt(sheet, 'A1')?.value).toBe('plain');
    expect(cellAt(sheet, 'B1')?.value).toBe('bold plain');
    expect(cellAt(sheet, 'C1')?.value).toBe('日本語');
    expect(cellAt(sheet, 'A2')).toMatchObject({ value: 1234.5, type: 'number' });
    expect(sheet.styles?.['1-0']?.numberFormat).toBe('"$"#,##0.00');
    expect(cellAt(sheet, 'B2')?.formula).toBe('A2*2');
    expect(cellAt(sheet, 'C2')).toMatchObject({ value: 'xy', formula: '"x"&"y"' });
    expect(sheet.columnWidths?.[0]).toBe(23.5);
  });

  it('ignores the declared encoding of a part', () => {
    const parts = excelLike();
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace('encoding="UTF-8"', 'encoding="utf-16"')
    );
    expect(cellAt(ExcelBridge.read(zipSync(parts)).sheets[0], 'A2')?.value).toBe(1234.5);
  });

  it('rejects a part stored as UTF-16 instead of reading it as empty', () => {
    const parts = excelLike();
    const xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
    const utf16 = new Uint8Array(2 + xml.length * 2);
    utf16.set([0xff, 0xfe]);
    for (let i = 0; i < xml.length; i++) {
      utf16[2 + 2 * i] = xml.charCodeAt(i) & 255;
      utf16[3 + 2 * i] = xml.charCodeAt(i) >> 8;
    }
    parts['xl/worksheets/sheet1.xml'] = utf16;
    expect(() => ExcelBridge.read(zipSync(parts))).toThrow(
      /xl\/worksheets\/sheet1\.xml: Malformed XML: text outside the root element/
    );
  });
});
