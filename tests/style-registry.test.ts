import { describe, expect, it } from 'vitest';
import { StyleManager } from '../src/core/style-manager';
import { generateStylesXml } from '../src/core/xml-templates';

const entries = (xml: string, tag: string): string[] =>
  xml.match(new RegExp(`<${tag}[ >/]`, 'g')) ?? [];

describe('style registry ids', () => {
  it('numbers entries in first-use order after the defaults', () => {
    const registry = new StyleManager();
    expect(registry.getStyleId({ bold: true })).toBe(1);
    expect(registry.getStyleId({ italic: true })).toBe(2);
    expect(registry.getStyleId({ bold: true })).toBe(1);
    expect(registry.getFontsCount()).toBe(3);
    expect(registry.getFillsCount()).toBe(2);
    expect(registry.getBordersCount()).toBe(1);
    expect(registry.getCellXfsCount()).toBe(3);
  });

  it('gives one id to equal styles whatever the property order', () => {
    const registry = new StyleManager();
    const first = registry.getStyleId({ bold: true, background: '#FF0000' });
    expect(registry.getStyleId({ background: '#FF0000', bold: true })).toBe(first);
    expect(registry.getCellXfsCount()).toBe(2);
  });

  it('treats a falsy flag like an absent one', () => {
    const registry = new StyleManager();
    expect(registry.getStyleId({})).toBe(0);
    expect(registry.getStyleId({ bold: false, border: false, background: '' })).toBe(0);
    expect(registry.getFontsCount()).toBe(1);
    expect(registry.getBordersCount()).toBe(1);
  });

  it('shares one font and one fill between spellings of a colour', () => {
    const registry = new StyleManager();
    registry.getStyleId({ color: '#f00', background: '#0f0' });
    registry.getStyleId({ color: '#FF0000', background: '#00FF00', bold: true });
    expect(registry.getFontsCount()).toBe(3);
    expect(registry.getFillsCount()).toBe(3);
    const fonts = registry.generateFontsXml();
    expect(fonts.match(/FFFF0000/g)).toHaveLength(2);
  });

  it('shares an entry between a default font spelled out and left out', () => {
    const registry = new StyleManager();
    const plain = registry.getStyleId({ bold: true });
    const spelled = registry.getStyleId({ bold: true, fontName: 'Calibri', fontSize: 11 });
    expect(spelled).toBe(plain);
    expect(registry.getFontsCount()).toBe(2);
  });

  it('numbers custom formats from 164 and reuses a repeated code', () => {
    const registry = new StyleManager();
    registry.getStyleId({ numberFormat: '0.00' });
    registry.getStyleId({ numberFormat: '0.0%' });
    registry.getStyleId({ numberFormat: '0.00', bold: true });
    expect(registry.getNumFmtsCount()).toBe(2);
    expect(registry.generateNumFmtsXml()).toBe(
      '    <numFmt numFmtId="164" formatCode="0.00"/>\n    <numFmt numFmtId="165" formatCode="0.0%"/>'
    );
    expect(registry.generateCellXfsXml()).toContain('numFmtId="165"');
  });

  it('writes the date style once, after the styles registered before it', () => {
    const registry = new StyleManager();
    registry.getStyleId({ bold: true });
    const date = registry.getDateStyleId();
    expect(date).toBe(2);
    expect(registry.getDateStyleId()).toBe(date);
    expect(registry.generateCellXfsXml().split('\n')[date]).toBe(
      '    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    );
  });

  it('shares a conditional-format entry between equal rules', () => {
    const registry = new StyleManager();
    const first = registry.getDxfId({ bold: true, color: '#9c0006' });
    expect(registry.getDxfId({ color: '#9C0006', bold: true })).toBe(first);
    expect(registry.getDxfId({ bold: true, italic: false, color: '#9C0006' })).toBe(first);
    expect(registry.getDxfId({ background: '#FFC7CE' })).toBe(first + 1);
    expect(registry.getDxfsCount()).toBe(2);
  });

  it('keeps the counts equal to the entries it writes', () => {
    const registry = new StyleManager();
    registry.getStyleId({ bold: true, background: '#FF0000', border: true, numberFormat: '0.0' });
    registry.getStyleId({ align: 'center', wrapText: true });
    registry.getDxfId({ italic: true });
    expect(entries(registry.generateFontsXml(), 'font')).toHaveLength(registry.getFontsCount());
    expect(entries(registry.generateFillsXml(), 'fill')).toHaveLength(registry.getFillsCount());
    expect(entries(registry.generateBordersXml(), 'border')).toHaveLength(
      registry.getBordersCount()
    );
    expect(entries(registry.generateCellXfsXml(), 'xf')).toHaveLength(registry.getCellXfsCount());
    expect(entries(registry.generateDxfsXml(), 'dxf')).toHaveLength(registry.getDxfsCount());
  });
});

describe('style registry validation', () => {
  it('rejects an invalid colour when the style is registered', () => {
    expect(() => new StyleManager().getStyleId({ color: 'red' })).toThrow(/Invalid colour/);
    expect(() => new StyleManager().getStyleId({ background: 'blue' })).toThrow(/Invalid colour/);
    expect(() => new StyleManager().getDxfId({ color: 'red' })).toThrow(/Invalid colour/);
  });
});

describe('generateStylesXml', () => {
  it('writes the same markup without a registry as with an empty one', () => {
    expect(generateStylesXml()).toBe(generateStylesXml(new StyleManager()));
  });
});
