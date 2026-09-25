import { describe, it, expect } from 'vitest';
import { unzipSync, strFromU8, zipSync, strToU8 } from 'fflate';
import { ExcelWriter, ExcelBridge, Workbook, hyperlink } from '../src';
import type { CellStyle, ExcelData, Hyperlink } from '../src';

const RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const unzip = (buffer: Uint8Array): Record<string, string> =>
  Object.fromEntries(
    Object.entries(unzipSync(buffer)).map(([name, bytes]) => [name, strFromU8(bytes)])
  );

const report = (hyperlinks: Hyperlink[], extra: Partial<ExcelData> = {}): ExcelData => ({
  data: [
    ['Name', 'Docs', 'Contact', 'Jump'],
    ['excel-bridge', 'Read the docs', 'Email us', 'Go to Q1'],
    ['Other', 'x', 'y', 'z'],
  ],
  hyperlinks,
  ...extra,
});

const write = (...sheets: ExcelData[]) => new ExcelWriter().createWorkbookBuffer(sheets);

const readStyles = (styles: Record<string, CellStyle>, links: Hyperlink[]) =>
  ExcelBridge.read(write({ ...report(links), styles })).sheets[0].styles ?? {};

describe('Hyperlinks', () => {
  it('writes external links through sheet relationships', () => {
    const files = unzip(
      write(report([{ range: 'B2', url: 'https://example.com/docs?a=1&b=2', tooltip: 'Docs' }]))
    );
    const sheet = files['xl/worksheets/sheet1.xml'];

    expect(sheet).toContain(`xmlns:r="${RELS_NS}"`);
    expect(sheet).toContain('<hyperlink ref="B2" r:id="rId1" tooltip="Docs"/>');
    expect(files['xl/worksheets/_rels/sheet1.xml.rels']).toContain(
      `<Relationship Id="rId1" Type="${RELS_NS}/hyperlink" Target="https://example.com/docs?a=1&amp;b=2" TargetMode="External"/>`
    );
  });

  it('leaves content types unchanged', () => {
    const withLinks = unzip(write(report([{ range: 'B2', url: 'https://example.com' }])));
    const withoutLinks = unzip(write(report([])));

    expect(withLinks['[Content_Types].xml']).toBe(withoutLinks['[Content_Types].xml']);
  });

  it('writes internal links without relationships', () => {
    const files = unzip(
      write(report([hyperlink.internal('D2', 'Q1 Sales', 'A1', { display: 'Q1' })]))
    );
    const sheet = files['xl/worksheets/sheet1.xml'];

    expect(sheet).toContain(`<hyperlink ref="D2" location="'Q1 Sales'!A1" display="Q1"/>`);
    expect(sheet).not.toContain('xmlns:r=');
    expect(files['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined();
  });

  it('builds links with the helpers', () => {
    expect(hyperlink.url('B2', 'https://excel-bridge.dev', { tooltip: 'Home' })).toEqual({
      range: 'B2',
      url: 'https://excel-bridge.dev',
      tooltip: 'Home',
    });
    expect(hyperlink.email('C2', 'team@example.com', { subject: 'Hi & bye' })).toEqual({
      range: 'C2',
      url: 'mailto:team@example.com?subject=Hi%20%26%20bye',
    });
    expect(hyperlink.internal('D2', "Bob's Q1", 'b3')).toEqual({
      range: 'D2',
      location: "'Bob''s Q1'!B3",
    });
  });

  it('splits the fragment on write and joins it on read', () => {
    const buffer = write(report([{ range: 'B2', url: 'https://example.com/docs#install' }]));
    const files = unzip(buffer);

    expect(files['xl/worksheets/sheet1.xml']).toContain(
      '<hyperlink ref="B2" r:id="rId1" location="install"/>'
    );
    expect(files['xl/worksheets/_rels/sheet1.xml.rels']).toContain(
      'Target="https://example.com/docs"'
    );
    expect(ExcelBridge.read(buffer).sheets[0].hyperlinks).toEqual([
      { range: 'B2', url: 'https://example.com/docs#install' },
    ]);
  });

  it('round-trips links and omits the field when absent', () => {
    const links: Hyperlink[] = [
      { range: 'B2', url: 'https://example.com/docs' },
      { range: 'C2', url: 'mailto:team@example.com' },
      { range: 'D2', location: "'Q1'!A1", tooltip: 'Jump', display: 'Q1' },
    ];

    expect(ExcelBridge.read(write(report(links))).sheets[0].hyperlinks).toEqual(links);
    expect(ExcelBridge.read(write({ data: [['x']] })).sheets[0]).not.toHaveProperty('hyperlinks');
  });

  it('escapes attribute values and reads them back', () => {
    const links: Hyperlink[] = [
      {
        range: 'B2',
        url: 'https://example.com/?q=a&b=c',
        tooltip: `" & < > '`,
        display: 'Say "hi" & <go>',
      },
      { range: 'D2', location: `'It''s "Q1"'!A1` },
    ];

    expect(ExcelBridge.read(write(report(links))).sheets[0].hyperlinks).toEqual(links);
  });

  it('drops a leading # from locations', () => {
    const buffer = write(report([{ range: 'D2', location: "#'Q1'!A1" }]));

    expect(ExcelBridge.read(buffer).sheets[0].hyperlinks).toEqual([
      { range: 'D2', location: "'Q1'!A1" },
    ]);
  });

  it('keeps links to sheets named like numbers through a Workbook round-trip', () => {
    const buffer = write(
      report([hyperlink.internal('D2', '007')], {
        options: { name: 'Index', autoFilter: { range: 'A1:D3' } },
      }),
      { data: [['target']], options: { name: '007', autoFilter: { range: 'A1:A1' } } }
    );
    const saved = Workbook.fromBuffer(buffer).toBuffer();
    const parsed = ExcelBridge.read(saved);

    expect(parsed.sheets.map(sheet => sheet.name)).toEqual(['Index', '007']);
    expect(parsed.sheets[0].hyperlinks).toEqual([{ range: 'D2', location: "'007'!A1" }]);
    expect(unzip(saved)['xl/workbook.xml']).toContain(`'007'!$A$1</definedName>`);
  });

  it('keeps numeric-looking text as strings', () => {
    const links: Hyperlink[] = [
      { range: 'B2', url: 'https://example.com', tooltip: 'true', display: '007' },
      { range: 'D2', location: '1e3' },
    ];

    expect(ExcelBridge.read(write(report(links))).sheets[0].hyperlinks).toEqual(links);
  });

  it('restarts relationship ids per sheet and writes rels only where needed', () => {
    const files = unzip(
      write(
        report(
          [
            { range: 'B2', url: 'https://a.dev' },
            { range: 'C2', url: 'mailto:a@a.dev' },
          ],
          { options: { name: 'One' } }
        ),
        report([hyperlink.internal('D2', 'One')], { options: { name: 'Two' } }),
        report([{ range: 'B2', url: 'https://c.dev' }], { options: { name: 'Three' } })
      )
    );

    expect(files['xl/worksheets/_rels/sheet1.xml.rels']).toContain('Id="rId2"');
    expect(files['xl/worksheets/_rels/sheet2.xml.rels']).toBeUndefined();
    expect(files['xl/worksheets/sheet3.xml']).toContain('<hyperlink ref="B2" r:id="rId1"/>');
    expect(files['xl/worksheets/_rels/sheet3.xml.rels']).toContain('Target="https://c.dev"');
  });

  describe('link look', () => {
    it('styles the first cell of each link', () => {
      const styles = readStyles({}, [{ range: 'B2:C2', url: 'https://example.com' }]);

      expect(styles['1-1']).toMatchObject({ color: '#0563C1', underline: true });
      expect(styles['1-2']).toBeUndefined();
    });

    it('lets the cell style win', () => {
      const styles = readStyles(
        { '1-1': { bold: true }, '2-1': { color: '#FFFFFF' }, '2-2': { underline: false } },
        [
          { range: 'B2', url: 'https://a.dev' },
          { range: 'B3', url: 'https://b.dev' },
          { range: 'C3', url: 'https://c.dev' },
        ]
      );

      expect(styles['1-1']).toMatchObject({ bold: true, color: '#0563C1', underline: true });
      expect(styles['2-1']).toMatchObject({ color: '#FFFFFF', underline: true });
      expect(styles['2-2']).toMatchObject({ color: '#0563C1' });
      expect(styles['2-2'].underline).toBeUndefined();
    });

    it('treats properties set to undefined as unset', () => {
      const styles = readStyles({ '1-1': { bold: true, color: undefined, underline: undefined } }, [
        { range: 'B2', url: 'https://a.dev' },
      ]);

      expect(styles['1-1']).toMatchObject({ bold: true, color: '#0563C1', underline: true });
    });
  });

  describe('validation', () => {
    const longUrl = `https://a.dev/${'x'.repeat(2066)}`;

    it.each([
      [{ range: 'B2' }, 'Hyperlink at B2 must set exactly one of url or location'],
      [
        { range: 'B2', url: 'https://a.dev', location: 'Sheet1!A1' },
        'Hyperlink at B2 must set exactly one of url or location',
      ],
      [
        { range: 'B2', url: 'javascript:alert(1)' },
        'Unsupported hyperlink url at B2: javascript:alert(1) (use http, https or mailto; percent-encode spaces and quotes)',
      ],
      [{ range: 'B2', url: 'file:///etc/passwd' }, 'Unsupported hyperlink url at B2'],
      [{ range: 'B2', url: 'https://a.dev/\n' }, 'Unsupported hyperlink url at B2'],
      [{ range: 'B2', url: 'mailto:a@b.co?subject=Q1 report' }, 'Unsupported hyperlink url at B2'],
      [{ range: 'B2', url: 'https://a.dev/?q="x"' }, 'Unsupported hyperlink url at B2'],
      [{ range: 'B2', url: 'https://' }, 'Unsupported hyperlink url at B2'],
      [{ range: 'B2', location: '#' }, 'Hyperlink at B2 must set exactly one of url or location'],
      [{ range: 'B2', url: longUrl }, 'Hyperlink url length 2080 exceeds Excel limit (2079)'],
      [{ range: 'B2', location: 'Sheet1!A1\u0007' }, 'Invalid hyperlink location at B2'],
      [{ range: 'ZZZZ1', url: 'https://a.dev' }, 'Invalid range format: ZZZZ1'],
    ])('rejects %o', (link, message) => {
      expect(() => write(report([link as unknown as Hyperlink]))).toThrow(message);
    });

    it('rejects two links on the same cell', () => {
      expect(() =>
        write(
          report([
            { range: 'B2', url: 'https://a.dev' },
            { range: 'b2', url: 'https://b.dev' },
          ])
        )
      ).toThrow('Duplicate hyperlink at B2');
    });

    it('rejects more links than Excel allows', () => {
      const links = Array.from({ length: 65531 }, (_, index) => ({
        range: `A${index + 1}`,
        url: 'https://a.dev',
      }));

      expect(() => write({ data: [['x']], hyperlinks: links })).toThrow(
        'Hyperlink count 65531 exceeds Excel limit (65530)'
      );
    });

    it('rejects an invalid cell in the internal builder', () => {
      expect(() => hyperlink.internal('B2', 'Q1', 'A0')).toThrow('Invalid range format: A0');
    });
  });

  describe('files saved by other apps', () => {
    const excelStyleFile = (
      hyperlinksXml = '<hyperlink ref="A1" r:id="rId3" location="top"/><hyperlink ref="B1" location="Sheet1!A1"/><hyperlink ref="C1" r:id="rId9"/><hyperlink ref="D1" r:id="rId4"/>'
    ) => {
      const files = unzipSync(write({ data: [['a', 'b', 'c', 'd']] }));
      const sheet = strFromU8(files['xl/worksheets/sheet1.xml'])
        .replace('<worksheet xmlns=', `<worksheet xmlns:r="${RELS_NS}" xmlns=`)
        .replace('</worksheet>', `<hyperlinks>${hyperlinksXml}</hyperlinks></worksheet>`);
      files['xl/worksheets/sheet1.xml'] = strToU8(sheet);
      files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="${RELS_NS}/hyperlink" Target="https://example.com/page" TargetMode="External"/><Relationship Id="rId4" Type="${RELS_NS}/hyperlink" Target="file:///C:/secret.txt" TargetMode="External"/></Relationships>`
      );
      return zipSync(files);
    };

    it('reads every stored link and skips dangling ids', () => {
      expect(ExcelBridge.read(excelStyleFile()).sheets[0].hyperlinks).toEqual([
        { range: 'A1', url: 'https://example.com/page#top' },
        { range: 'B1', location: 'Sheet1!A1' },
        { range: 'D1', url: 'file:///C:/secret.txt' },
      ]);
    });

    it('drops links the writer refuses when loading into a Workbook', () => {
      const workbook = Workbook.fromBuffer(excelStyleFile());

      expect(workbook.getHyperlinks('Sheet1')).toEqual([
        { range: 'A1', url: 'https://example.com/page#top' },
        { range: 'B1', location: 'Sheet1!A1' },
      ]);
      expect(() => workbook.toBuffer()).not.toThrow();
    });

    it('keeps the first of several links on one range when loading into a Workbook', () => {
      const workbook = Workbook.fromBuffer(
        excelStyleFile(
          '<hyperlink ref="A1" location="Sheet1!B1"/><hyperlink ref="A1:A1" r:id="rId3"/><hyperlink ref="a1" location="Sheet1!C1"/>'
        )
      );

      expect(workbook.getHyperlinks('Sheet1')).toEqual([{ range: 'A1', location: 'Sheet1!B1' }]);
      expect(ExcelBridge.read(workbook.toBuffer()).sheets[0].hyperlinks).toEqual([
        { range: 'A1', location: 'Sheet1!B1' },
      ]);
    });
  });

  describe('Workbook', () => {
    it('keeps links through load, edit and save', () => {
      const workbook = Workbook.fromBuffer(
        write(report([{ range: 'B2', url: 'https://example.com' }], { options: { name: 'Links' } }))
      );
      workbook.setCellValue('Links', 1, 0, 'renamed');

      expect(ExcelBridge.read(workbook.toBuffer()).sheets[0].hyperlinks).toEqual([
        { range: 'B2', url: 'https://example.com' },
      ]);
    });

    it('sets, replaces, copies and removes links', () => {
      const workbook = Workbook.create();
      workbook.addSheet('Links', [['a', 'b']]);

      workbook.setHyperlink('Links', { range: 'a1', url: 'https://one.dev' });
      workbook.setHyperlink('Links', { range: 'A1', url: 'https://two.dev' });
      expect(workbook.getHyperlinks('Links')).toEqual([{ range: 'A1', url: 'https://two.dev' }]);

      workbook.getHyperlinks('Links')[0].range = 'Z9';
      expect(workbook.getHyperlinks('Links')[0].range).toBe('A1');

      workbook.removeHyperlink('Links', 'B1');
      workbook.removeHyperlink('Links', 'A1');
      expect(workbook.getHyperlinks('Links')).toEqual([]);
    });

    it('removes the link look together with the link', () => {
      const workbook = Workbook.fromBuffer(
        write({
          data: [['a', 'b']],
          styles: { '0-0': { bold: true } },
          hyperlinks: [{ range: 'A1', url: 'https://example.com' }],
          options: { name: 'Links' },
        })
      );
      expect(workbook.getCellStyle('Links', 0, 0)).toMatchObject({
        bold: true,
        color: '#0563C1',
        underline: true,
      });

      workbook.removeHyperlink('Links', 'A1');
      const style = workbook.getCellStyle('Links', 0, 0);

      expect(style).toMatchObject({ bold: true });
      expect(style?.underline).toBeUndefined();
      expect(style?.color).toBeUndefined();
    });

    it('validates at the call', () => {
      const workbook = Workbook.create();
      workbook.addSheet('Links');

      expect(() => workbook.setHyperlink('Links', { range: 'B2', url: 'ftp://x.dev' })).toThrow(
        'Unsupported hyperlink url at B2'
      );
      expect(() => workbook.setHyperlink('Missing', { range: 'B2', url: 'https://x.dev' })).toThrow(
        'Sheet "Missing" not found'
      );
    });
  });
});
