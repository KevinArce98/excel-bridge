import { describe, expect, it } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { strFromU8, unzipSync } from 'fflate';
import { parseXml } from '../src/reader/xml';
import { fixture } from './helpers/read';
import { randomXml } from './helpers/random-xml';

const oracle = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

type Tree = Record<string, unknown>;

const withoutProcessingInstructions = (node: unknown, top = true): unknown => {
  if (Array.isArray(node)) return node.map(item => withoutProcessingInstructions(item, false));
  if (node === null || typeof node !== 'object') return node;
  const kept: Tree = {};
  let removed = false;
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith('?')) removed = true;
    else kept[key] = withoutProcessingInstructions(value, false);
  }
  const keys = Object.keys(kept);
  if (!top && removed && keys.length === 0) return '';
  if (!top && removed && keys.length === 1 && keys[0] === '#text') return kept['#text'];
  return kept;
};

const decodeNumericReferences = (xml: string): string =>
  xml.replace(/&#(?:(\d+)|x([\da-fA-F]+));/g, (all, decimal?: string, hex?: string) => {
    const code = decimal ? parseInt(decimal, 10) : parseInt(hex as string, 16);
    if (code === 13) return all;
    const char = String.fromCodePoint(code);
    return (
      (
        { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' } as Record<
          string,
          string
        >
      )[char] ?? char
    );
  });

const withoutRootPrefix = (xml: string): string => {
  const bare = xml.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, '');
  const root = /<([A-Za-z_:\u0080-￿][^\s/>]*)/.exec(bare)?.[1] ?? '';
  const colon = root.indexOf(':');
  if (colon <= 0) return xml;
  const prefix = root.slice(0, colon + 1);
  return xml.split(`<${prefix}`).join('<').split(`</${prefix}`).join('</');
};

const expected = (xml: string): unknown => {
  const tree = withoutProcessingInstructions(
    oracle.parse(withoutRootPrefix(decodeNumericReferences(xml)))
  ) as Tree;
  delete tree['#text'];
  return tree;
};

describe('parseXml against fast-xml-parser', () => {
  it.each(['exceljs-report.xlsx', 'exceljs-text.xlsx', 'sheetjs-report.xlsx', 'hucre-report.xlsx'])(
    'builds the same tree for every XML part of %s',
    file => {
      for (const [path, content] of Object.entries(unzipSync(fixture(file)))) {
        if (!/\.(xml|rels)$/.test(path)) continue;
        const xml = strFromU8(content);
        expect(parseXml(xml), path).toEqual(expected(xml));
      }
    }
  );

  it('builds the same tree for 5000 generated documents', () => {
    for (let seed = 1; seed <= 5000; seed++) {
      const xml = randomXml(seed);
      if (/&#(13|x0*[dD]);/.test(xml)) continue;
      expect(parseXml(xml), `seed ${seed}`).toEqual(expected(xml));
    }
  });
});
