import { describe, expect, it } from 'vitest';
import { parseXml } from '../src/reader/xml';

const parse = (xml: string) => parseXml(xml);

describe('parseXml shape', () => {
  it('returns text for an element with only text', () => {
    expect(parse('<a><b>hi</b></a>')).toEqual({ a: { b: 'hi' } });
  });

  it('returns an empty string for an empty element', () => {
    expect(parse('<a><b/><c></c></a>')).toEqual({ a: { b: '', c: '' } });
  });

  it('returns attributes as strings and keeps empty values', () => {
    expect(parse('<a><b x="1" y=""/></a>')).toEqual({ a: { b: { x: '1', y: '' } } });
  });

  it('puts the text of an element with attributes under #text', () => {
    expect(parse('<a><b x="1">hi</b></a>')).toEqual({ a: { b: { x: '1', '#text': 'hi' } } });
  });

  it('turns repeated siblings into an array and keeps a single one as is', () => {
    expect(parse('<a><b>1</b><b>2</b><c>3</c></a>')).toEqual({ a: { b: ['1', '2'], c: '3' } });
    expect(parse('<a><b/><b/><b/></a>')).toEqual({ a: { b: ['', '', ''] } });
  });

  it('joins text around children under #text', () => {
    expect(parse('<a>x<b>1</b>y</a>')).toEqual({ a: { b: '1', '#text': 'xy' } });
  });

  it('keeps whitespace-only text', () => {
    expect(parse('<a><b> </b><c>\n</c></a>')).toEqual({ a: { b: ' ', c: '\n' } });
    expect(parse('<a>\n <b>1</b>\n</a>')).toEqual({ a: { b: '1', '#text': '\n \n' } });
    expect(parse('<a><t xml:space="preserve"> </t></a>')).toEqual({
      a: { t: { 'xml:space': 'preserve', '#text': ' ' } },
    });
  });

  it('returns an empty object for a document with no element', () => {
    expect(parse('')).toEqual({});
    expect(parse('<?xml version="1.0"?>')).toEqual({});
  });

  it('does not treat attribute and child names specially', () => {
    expect(parse('<a><hasOwnProperty>1</hasOwnProperty><toString>2</toString></a>')).toEqual({
      a: { hasOwnProperty: '1', toString: '2' },
    });
    expect(parse('<a><toString/><toString/></a>')).toEqual({ a: { toString: ['', ''] } });
  });
});

describe('parseXml text', () => {
  it('decodes the five predefined entities once', () => {
    expect(parse('<a>&lt;&gt;&amp;&quot;&apos;</a>')).toEqual({ a: `<>&"'` });
    expect(parse('<a>&amp;lt;</a>')).toEqual({ a: '&lt;' });
    expect(parse('<a x="&lt;&amp;&quot;&apos;"/>')).toEqual({ a: { x: `<&"'` } });
  });

  it('decodes numeric character references', () => {
    expect(parse('<a>&#233;&#xE9;&#x41;&#128512;&#10;&#9;</a>')).toEqual({ a: 'ééA😀\n\t' });
    expect(parse('<a x="&#65;&#x42;"/>')).toEqual({ a: { x: 'AB' } });
  });

  it('keeps a numeric reference to a carriage return', () => {
    expect(parse('<a>a&#13;b&#xD;&#10;c</a>')).toEqual({ a: 'a\rb\r\nc' });
  });

  it('leaves references that are not characters as they are', () => {
    expect(parse('<a>&#0;&#xD800;&#1114112;&#xZZ;&#;&amp</a>')).toEqual({
      a: '&#0;&#xD800;&#1114112;&#xZZ;&#;&amp',
    });
  });

  it('leaves a bare ampersand and unknown entities as they are', () => {
    expect(parse('<a>a & b &nbsp; &foo;</a>')).toEqual({ a: 'a & b &nbsp; &foo;' });
  });

  it('normalizes line ends to \\n in text, attributes and CDATA', () => {
    expect(parse('<a>x\r\ny\rz</a>')).toEqual({ a: 'x\ny\nz' });
    expect(parse('<a b="x\r\ny"/>')).toEqual({ a: { b: 'x\ny' } });
    expect(parse('<a><![CDATA[x\r\ny]]></a>')).toEqual({ a: 'x\ny' });
  });

  it('keeps tabs and newlines in attribute values', () => {
    expect(parse('<a b="x\ty\nz"/>')).toEqual({ a: { b: 'x\ty\nz' } });
  });

  it('reads CDATA as text without decoding it', () => {
    expect(parse('<a><![CDATA[x<y&amp;]]></a>')).toEqual({ a: 'x<y&amp;' });
    expect(parse('<a>p<![CDATA[q]]>r&amp;s</a>')).toEqual({ a: 'pqr&s' });
    expect(parse('<a><![CDATA[]]></a>')).toEqual({ a: '' });
  });

  it('skips comments and processing instructions', () => {
    expect(
      parse(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a><!-- c --><b>x<!-- c -->y</b><?pi x?></a>'
      )
    ).toEqual({
      a: { b: 'xy' },
    });
  });

  it('allows > in text and attribute values', () => {
    expect(parse('<a x="1>2">1 > 2 ]]</a>')).toEqual({ a: { x: '1>2', '#text': '1 > 2 ]]' } });
  });

  it('allows either quote and whitespace around =', () => {
    expect(parse(`<a x = '1' y\n=\t"2" />`)).toEqual({ a: { x: '1', y: '2' } });
  });

  it('skips a leading BOM and whitespace outside the root', () => {
    expect(parse('﻿  \n<a>1</a>\r\n')).toEqual({ a: '1' });
  });
});

describe('parseXml namespaces', () => {
  it('drops the prefix of the root element from every element name', () => {
    const xml =
      '<x:worksheet xmlns:x="u"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>5</x:v></x:c></x:row></x:sheetData></x:worksheet>';
    expect(parse(xml)).toEqual({
      worksheet: { sheetData: { row: { r: '1', c: { r: 'A1', v: '5' } } }, 'xmlns:x': 'u' },
    });
  });

  it('keeps other prefixes and attribute prefixes', () => {
    const xml =
      '<cp:coreProperties xmlns:cp="a" xmlns:dc="b"><dc:creator>me</dc:creator><cp:keywords/></cp:coreProperties>';
    expect(parse(xml)).toEqual({
      coreProperties: { 'dc:creator': 'me', keywords: '', 'xmlns:cp': 'a', 'xmlns:dc': 'b' },
    });
    expect(parse('<sheet r:id="rId1" xml:space="preserve"/>')).toEqual({
      sheet: { 'r:id': 'rId1', 'xml:space': 'preserve' },
    });
  });

  it('keeps the prefix of elements from other namespaces', () => {
    expect(parse('<worksheet><x14:cfRule/><mc:AlternateContent/></worksheet>')).toEqual({
      worksheet: { 'x14:cfRule': '', 'mc:AlternateContent': '' },
    });
  });
});

describe('parseXml malformed input', () => {
  it.each([
    ['an unclosed element', '<a><b>1</b>', /<a> is not closed/],
    ['a closing tag with no element', '</a>', /unexpected closing tag at position 0$/],
    ['a mismatched closing tag', '<a><b>1</c></a>', /unexpected closing tag, expected <\/b>/],
    ['a longer closing name', '<a></ab>', /unexpected closing tag/],
    ['a closing tag that is not terminated', '<a></a', /closing tag is not terminated/],
    ['an unquoted attribute', '<a x=1/>', /attribute x is not quoted/],
    ['an attribute without a value', '<a x/>', /attribute x has no value/],
    ['an unterminated attribute', '<a x="1/>', /attribute x is not terminated/],
    ['attributes with no space between them', '<a x="1"y="2"/>', /invalid attribute name/],
    ['a stray slash', '<a / >', /stray "\/" in tag/],
    ['a tag that starts with a digit', '<1a/>', /invalid tag name/],
    ['a lone <', '<a>1 < 2</a>', /invalid tag name/],
    ['a tag name with a quote', '<a"b/>', /invalid attribute name/],
    ['text outside the root', 'x<a/>', /text outside the root element/],
    ['text after the root', '<a/>x', /text outside the root element/],
    ['CDATA outside the root', '<![CDATA[x]]><a/>', /CDATA outside the root element/],
    ['an unterminated comment', '<a><!-- x</a>', /comment is not terminated/],
    ['an unterminated CDATA', '<a><![CDATA[x</a>', /CDATA section is not terminated/],
    [
      'an unterminated processing instruction',
      '<a><?pi x</a>',
      /processing instruction is not terminated/,
    ],
    ['a DOCTYPE', '<!DOCTYPE a><a/>', /DOCTYPE and other declarations are not supported/],
    ['an element named __proto__', '<a><__proto__>1</__proto__></a>', /reserved name/],
    ['an attribute named __proto__', '<a __proto__="1"/>', /reserved name/],
    [
      'an element named __proto__ behind the prefix of the root',
      '<x:a><x:__proto__><b>1</b></x:__proto__></x:a>',
      /reserved name/,
    ],
    [
      'an empty element named __proto__ behind the prefix of the root',
      '<x:a><x:__proto__/></x:a>',
      /reserved name/,
    ],
    ['a root named __proto__ behind its own prefix', '<x:__proto__/>', /reserved name/],
  ])('throws for %s', (_label, xml, message) => {
    expect(() => parse(xml)).toThrow(message);
  });

  it('names the position of the error', () => {
    expect(() => parse('<a>\n<b>1</c></a>')).toThrow(/at position 8$/);
  });

  it('puts an unclosed element at the end of the input', () => {
    expect(() => parse('<a><b>1</b>')).toThrow(/<a> is not closed at position 11$/);
    expect(() => parse('<a>text')).toThrow(/<a> is not closed at position 7$/);
    expect(() => parse('<a><b>')).toThrow(/<b> is not closed at position 6$/);
  });

  it('does not name an element when a closing tag has nothing to close', () => {
    expect(() => parse('<a/></a>')).toThrow(/Malformed XML: unexpected closing tag at position 4$/);
  });

  it('does not expand entities declared in a DOCTYPE', () => {
    const bomb = `<!DOCTYPE a [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">${'<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">'}]><a>&c;</a>`;
    expect(() => parse(bomb)).toThrow(/DOCTYPE/);
    expect(() => parse('<?xml version="1.0"?><!DOCTYPE a [<!ENTITY e "x">]><a>&e;</a>')).toThrow(
      /DOCTYPE/
    );
  });

  it('accepts 101 levels of nesting and rejects 102', () => {
    const nested = (levels: number) => '<a>'.repeat(levels) + '</a>'.repeat(levels);
    expect(() => parse(nested(101))).not.toThrow();
    expect(() => parse(nested(102))).toThrow(/nested too deeply/);
  });

  it('rejects very deep nesting without recursing', () => {
    expect(() => parse('<a>'.repeat(2_000_000))).toThrow(/nested too deeply/);
    expect(() => parse('</a>'.repeat(1_000_000))).toThrow(/unexpected closing tag/);
  });

  it('stays linear on hostile input', () => {
    const inputs = [
      `<a>${'&'.repeat(3_000_000)}</a>`,
      `<a>${'&#'.repeat(1_500_000)}</a>`,
      `<a>${'&#x1'.repeat(1_000_000)}</a>`,
      `<a>${'<!-- -->'.repeat(500_000)}</a>`,
      `<a>${'<?p?>'.repeat(500_000)}</a>`,
      `<a>${'<![CDATA[]]>'.repeat(300_000)}</a>`,
      `<a>${'<b/>'.repeat(500_000)}</a>`,
      `<a>${'x'.repeat(5_000_000)}</a>`,
      `<a ${Array.from({ length: 100_000 }, (_, i) => `a${i}="1"`).join(' ')}/>`,
    ];
    for (const xml of inputs) {
      const started = performance.now();
      parse(xml);
      expect(performance.now() - started).toBeLessThan(3000);
    }
    for (const xml of ['<a>' + '<!--'.repeat(2_000_000), '<a x="' + "'".repeat(5_000_000)]) {
      const started = performance.now();
      expect(() => parse(xml)).toThrow();
      expect(performance.now() - started).toBeLessThan(3000);
    }
  });
});
