import { describe, it, expect } from 'vitest';
import { StyleManager } from '../src/core/style-manager';
import { generateStylesXml } from '../src/core/xml-templates';
import { StyleManager as LegacyStyleManager } from './helpers/legacy-style-manager';
import type { CellStyle, ConditionalFormatStyle } from '../src/core/types';

const mulberry32 = (seed: number) => () => {
  seed |= 0;
  seed = (seed + 0x6d2b79f5) | 0;
  let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
};

type Pool = Record<string, unknown[]>;

const WILD: Pool = {
  background: [undefined, '', '#FF0000', '#f00', '#ff0000', '#FFFF0000', '#00FF00'],
  bold: [undefined, true, false],
  italic: [undefined, true, false],
  underline: [undefined, true, false],
  border: [undefined, true, false],
  color: [undefined, '', '#0563C1', '#abc', '#AABBCC', '#112233', '#FF112233'],
  fontSize: [undefined, 0, 11, 14],
  fontName: [undefined, '', 'Calibri', 'Arial', 'A&B "x" <y>'],
  align: [undefined, '', 'left', 'center', 'right'],
  verticalAlign: [undefined, '', 'top', 'middle', 'bottom'],
  wrapText: [undefined, true, false],
  numberFormat: [undefined, '', '0.00', '#,##0', '0.0%', '"x" & <y>'],
};

const CLEAN: Pool = {
  background: [undefined, '#FF0000', '#00FF00'],
  bold: [undefined, true],
  italic: [undefined, true],
  underline: [undefined, true],
  border: [undefined, true],
  color: [undefined, '#0563C1', '#112233'],
  fontSize: [undefined, 14],
  fontName: [undefined, 'Arial', 'A&B "x" <y>'],
  align: [undefined, 'left', 'center', 'right'],
  verticalAlign: [undefined, 'top', 'middle', 'bottom'],
  wrapText: [undefined, true],
  numberFormat: [undefined, '0.00', '#,##0', '0.0%', '"x" & <y>'],
};

const DXF_WILD: Pool = {
  bold: [undefined, true, false],
  italic: [undefined, true, false],
  color: [undefined, '', '#9C0006', '#9c0006', '#f00'],
  background: [undefined, '', '#FFC7CE', '#ffc7ce'],
};
const DXF_CLEAN: Pool = {
  bold: [undefined, true],
  italic: [undefined, true],
  color: [undefined, '#9C0006', '#006100'],
  background: [undefined, '#FFC7CE', '#C6EFCE'],
};

const pick = <T>(random: () => number, items: T[]): T => items[Math.floor(random() * items.length)];

const draw = (random: () => number, pool: Pool, shuffleKeys: boolean): Record<string, unknown> => {
  const keys = Object.keys(pool);
  if (shuffleKeys) keys.sort(() => random() - 0.5);
  const out: Record<string, unknown> = {};
  for (const key of keys) {
    const value = pick(random, pool[key]);
    if (value !== undefined || random() < 0.2) out[key] = value;
  }
  return out;
};

type Call =
  | { kind: 'style'; style: CellStyle }
  | { kind: 'date' }
  | { kind: 'dxf'; style: ConditionalFormatStyle };

const sequence = (seed: number, pool: Pool, dxfPool: Pool, requireNonEmpty: boolean): Call[] => {
  const random = mulberry32(seed);
  const length = 1 + Math.floor(random() * 40);
  const calls: Call[] = [];
  for (let i = 0; i < length; i++) {
    const roll = random();
    if (roll < 0.06) calls.push({ kind: 'date' });
    else if (roll < 0.14) {
      const style = draw(random, dxfPool, false) as ConditionalFormatStyle;
      if (!requireNonEmpty || Object.keys(style).length > 0) calls.push({ kind: 'dxf', style });
    } else {
      const style = draw(random, pool, requireNonEmpty) as CellStyle;
      if (!requireNonEmpty || Object.keys(style).length > 0) calls.push({ kind: 'style', style });
    }
  }
  return calls;
};

type Registry = LegacyStyleManager | StyleManager;

const run = (registry: Registry, calls: Call[]): number[] =>
  calls.map(call =>
    call.kind === 'style'
      ? registry.getStyleId(call.style)
      : call.kind === 'date'
        ? registry.getDateStyleId()
        : registry.getDxfId(call.style)
  );

const parts = (registry: Registry) => ({
  fonts: registry.generateFontsXml(),
  fills: registry.generateFillsXml(),
  borders: registry.generateBordersXml(),
  cellXfs: registry.generateCellXfsXml(),
  numFmts: registry.generateNumFmtsXml(),
  dxfs: registry.generateDxfsXml(),
  counts: [
    registry.getFontsCount(),
    registry.getFillsCount(),
    registry.getBordersCount(),
    registry.getCellXfsCount(),
    registry.getNumFmtsCount(),
    registry.getDxfsCount(),
  ],
  styles: generateStylesXml(registry as unknown as StyleManager),
});

const splitEntries = (xml: string, opening: string): string[] =>
  xml === '' ? [] : xml.split(new RegExp(`\\n(?=${opening})`));

const resolveStyle = (registry: Registry, id: number) => {
  const p = parts(registry);
  const xf = splitEntries(p.cellXfs, '    <xf ')[id];
  const read = (name: string) => Number(new RegExp(`${name}="(\\d+)"`).exec(xf)![1]);
  const numFmtId = read('numFmtId');
  const numFmtLine = splitEntries(p.numFmts, '    <numFmt ').find(line =>
    line.includes(`numFmtId="${numFmtId}"`)
  );
  return {
    font: splitEntries(p.fonts, '    <font>')[read('fontId')],
    fill: splitEntries(p.fills, '    <fill>')[read('fillId')],
    border: splitEntries(p.borders, '    <border>')[read('borderId')],
    numFmt: numFmtId < 164 ? numFmtId : numFmtLine,
    rest: xf.replace(/numFmtId="\d+" fontId="\d+" fillId="\d+" borderId="\d+" /, ''),
  };
};

const resolveDxf = (registry: Registry, id: number) =>
  splitEntries(parts(registry).dxfs, '    <dxf>')[id];

const SEEDS = 600;

const compareExact = (pool: Pool, dxfPool: Pool, requireNonEmpty: boolean) => {
  let divergent = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const calls = sequence(seed, pool, dxfPool, requireNonEmpty);
    const legacy = new LegacyStyleManager();
    const next = new StyleManager();
    const legacyIds = run(legacy, calls);
    const nextIds = run(next, calls);
    const same =
      JSON.stringify(legacyIds) === JSON.stringify(nextIds) &&
      JSON.stringify(parts(legacy)) === JSON.stringify(parts(next));
    if (!same) {
      divergent++;
    }
  }
  return divergent;
};

const compareSemantic = (pool: Pool, dxfPool: Pool) => {
  let mismatches = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const calls = sequence(seed, pool, dxfPool, false);
    const legacy = new LegacyStyleManager();
    const next = new StyleManager();
    const legacyIds = run(legacy, calls);
    const nextIds = run(next, calls);
    calls.forEach((call, index) => {
      if (call.kind === 'dxf') {
        if (resolveDxf(legacy, legacyIds[index]) !== resolveDxf(next, nextIds[index])) mismatches++;
      } else if (
        JSON.stringify(resolveStyle(legacy, legacyIds[index])) !==
        JSON.stringify(resolveStyle(next, nextIds[index]))
      ) {
        mismatches++;
      }
    });
  }
  return mismatches;
};

describe('style registry against the 1.5.0 implementation', () => {
  it('writes identical ids and bytes for styles without redundant entries', () => {
    expect(compareExact(CLEAN, DXF_CLEAN, true)).toBe(0);
  });

  it('resolves every style, redundant or not, to the same effective formatting', () => {
    const mismatches = compareSemantic(WILD, DXF_WILD);
    expect(mismatches).toBe(0);
  });

  it('resolves every style without redundant entries to the same effective formatting', () => {
    expect(compareSemantic(CLEAN, DXF_CLEAN)).toBe(0);
  });
});
