import { CellStyle } from './types';

export interface CellAlignment {
  horizontal?: 'left' | 'center' | 'right';
  vertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
}

export interface ExcelStyle {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  alignment?: CellAlignment;
}

export interface Font {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  size?: number;
  name?: string;
}

export interface Fill {
  fgColor?: string;
  bgColor?: string;
  patternType?: string;
}

export interface Border {
  left?: boolean;
  right?: boolean;
  top?: boolean;
  bottom?: boolean;
  color?: string;
}

export class StyleManager {
  private fonts: Font[] = [];
  private fills: Fill[] = [];
  private borders: Border[] = [];
  private cellXfs: ExcelStyle[] = [];
  private styleMap: Map<string, number> = new Map();
  // Custom number formats: format code -> numFmtId (>= 164)
  private numFmts: Map<string, number> = new Map();
  private nextNumFmtId = 164;

  constructor() {
    // Add default font
    this.fonts.push({ size: 11, name: 'Calibri' });

    // Add default fills (required by Excel)
    this.fills.push({ patternType: 'none' });
    this.fills.push({ patternType: 'gray125' });

    // Add default border
    this.borders.push({});

    // Add default cellXf
    this.cellXfs.push({
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 0,
    });
  }

  getStyleId(style: CellStyle): number {
    const hash = this.hashStyle(style);

    if (this.styleMap.has(hash)) {
      return this.styleMap.get(hash)!;
    }

    const fontId = this.addFont({
      bold: style.bold,
      italic: style.italic,
      underline: style.underline,
      color: style.color,
      size: style.fontSize,
      name: style.fontName,
    });

    const fillId = this.addFill({
      fgColor: style.background,
      patternType: style.background ? 'solid' : 'none',
    });

    const borderId = this.addBorder({
      left: style.border,
      right: style.border,
      top: style.border,
      bottom: style.border,
    });

    const numFmtId = style.numberFormat ? this.addNumFmt(style.numberFormat) : 0;

    const hasAlignment = !!(style.align || style.verticalAlign || style.wrapText);
    const alignment: CellAlignment | undefined = hasAlignment
      ? { horizontal: style.align, vertical: style.verticalAlign, wrapText: style.wrapText }
      : undefined;

    const hasFont =
      style.bold ||
      style.italic ||
      style.underline ||
      !!style.color ||
      !!style.fontSize ||
      !!style.fontName;

    const cellXf: ExcelStyle = {
      fontId,
      fillId,
      borderId,
      numFmtId,
      applyFont: hasFont,
      applyFill: !!style.background,
      applyBorder: !!style.border,
      applyNumberFormat: !!style.numberFormat,
      alignment,
    };

    this.cellXfs.push(cellXf);
    const styleId = this.cellXfs.length - 1;
    this.styleMap.set(hash, styleId);

    return styleId;
  }

  private addNumFmt(code: string): number {
    const existing = this.numFmts.get(code);
    if (existing !== undefined) {
      return existing;
    }
    const id = this.nextNumFmtId++;
    this.numFmts.set(code, id);
    return id;
  }

  getDateStyleId(): number {
    const hash = 'DATE_FORMAT_14';

    if (this.styleMap.has(hash)) {
      return this.styleMap.get(hash)!;
    }

    const cellXf: ExcelStyle = {
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 14, // Standard date format
      applyNumberFormat: true,
    };

    this.cellXfs.push(cellXf);
    const styleId = this.cellXfs.length - 1;
    this.styleMap.set(hash, styleId);

    return styleId;
  }

  private hashStyle(style: CellStyle): string {
    return JSON.stringify({
      bg: style.background || '',
      bold: style.bold || false,
      italic: style.italic || false,
      underline: style.underline || false,
      border: style.border || false,
      color: style.color || '',
      fontSize: style.fontSize || 0,
      fontName: style.fontName || '',
      align: style.align || '',
      valign: style.verticalAlign || '',
      wrap: style.wrapText || false,
      numFmt: style.numberFormat || '',
    });
  }

  private addFont(font: Font): number {
    const normalized: Font = {
      size: font.size || 11,
      name: font.name || 'Calibri',
      bold: font.bold,
      italic: font.italic,
      underline: font.underline,
      color: font.color,
    };

    const existing = this.fonts.findIndex(
      f =>
        f.bold === normalized.bold &&
        f.italic === normalized.italic &&
        f.underline === normalized.underline &&
        f.color === normalized.color &&
        f.size === normalized.size &&
        f.name === normalized.name
    );

    if (existing !== -1) {
      return existing;
    }

    this.fonts.push(normalized);

    return this.fonts.length - 1;
  }

  private addFill(fill: Fill): number {
    const existing = this.fills.findIndex(
      f => f.fgColor === fill.fgColor && f.patternType === fill.patternType
    );

    if (existing !== -1) {
      return existing;
    }

    this.fills.push(fill);
    return this.fills.length - 1;
  }

  private addBorder(border: Border): number {
    const existing = this.borders.findIndex(
      b =>
        b.left === border.left &&
        b.right === border.right &&
        b.top === border.top &&
        b.bottom === border.bottom
    );

    if (existing !== -1) {
      return existing;
    }

    this.borders.push(border);
    return this.borders.length - 1;
  }

  generateFontsXml(): string {
    return this.fonts
      .map(font => {
        let xml = '    <font>';
        if (font.bold) xml += '\n      <b/>';
        if (font.italic) xml += '\n      <i/>';
        if (font.underline) xml += '\n      <u/>';
        if (font.size) xml += `\n      <sz val="${font.size}"/>`;
        if (font.color) xml += `\n      <color rgb="${this.normalizeColor(font.color)}"/>`;
        if (font.name) xml += `\n      <name val="${font.name}"/>`;
        xml += '\n    </font>';
        return xml;
      })
      .join('\n');
  }

  generateFillsXml(): string {
    return this.fills
      .map(fill => {
        if (fill.patternType === 'none' || fill.patternType === 'gray125') {
          return `    <fill>\n      <patternFill patternType="${fill.patternType}"/>\n    </fill>`;
        }

        let xml = '    <fill>\n      <patternFill patternType="solid">';
        if (fill.fgColor) {
          xml += `\n        <fgColor rgb="${this.normalizeColor(fill.fgColor)}"/>`;
        }
        xml += '\n      </patternFill>\n    </fill>';
        return xml;
      })
      .join('\n');
  }

  generateBordersXml(): string {
    return this.borders
      .map(border => {
        let xml = '    <border>';
        xml += this.generateBorderSide('left', border.left);
        xml += this.generateBorderSide('right', border.right);
        xml += this.generateBorderSide('top', border.top);
        xml += this.generateBorderSide('bottom', border.bottom);
        xml += '\n      <diagonal/>';
        xml += '\n    </border>';
        return xml;
      })
      .join('\n');
  }

  private generateBorderSide(side: string, enabled?: boolean): string {
    if (!enabled) {
      return `<${side}/>`;
    }
    return `<${side} style="thin"><color rgb="FF000000"/></${side}>`;
  }

  generateCellXfsXml(): string {
    return this.cellXfs
      .map(xf => {
        let xml = `    <xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"`;
        if (xf.applyFont) xml += ' applyFont="1"';
        if (xf.applyFill) xml += ' applyFill="1"';
        if (xf.applyBorder) xml += ' applyBorder="1"';
        if (xf.applyNumberFormat) xml += ' applyNumberFormat="1"';
        if (xf.alignment) xml += ' applyAlignment="1"';

        if (xf.alignment) {
          const { horizontal, vertical, wrapText } = xf.alignment;
          // Excel uses "center" for vertical middle alignment.
          const verticalValue = vertical === 'middle' ? 'center' : vertical;
          let alignXml = '      <alignment';
          if (horizontal) alignXml += ` horizontal="${horizontal}"`;
          if (verticalValue) alignXml += ` vertical="${verticalValue}"`;
          if (wrapText) alignXml += ' wrapText="1"';
          alignXml += '/>';
          xml += `>\n${alignXml}\n    </xf>`;
        } else {
          xml += '/>';
        }
        return xml;
      })
      .join('\n');
  }

  generateNumFmtsXml(): string {
    return Array.from(this.numFmts.entries())
      .map(([code, id]) => `    <numFmt numFmtId="${id}" formatCode="${this.escapeAttr(code)}"/>`)
      .join('\n');
  }

  getNumFmtsCount(): number {
    return this.numFmts.size;
  }

  private escapeAttr(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  private normalizeColor(color: string): string {
    // Remove # if present
    let normalized = color.replace('#', '');

    // If 3-digit hex, expand to 6-digit
    if (normalized.length === 3) {
      normalized = normalized
        .split('')
        .map(c => c + c)
        .join('');
    }

    // Add alpha channel if not present (FF = opaque)
    if (normalized.length === 6) {
      normalized = 'FF' + normalized;
    }

    return normalized.toUpperCase();
  }

  getFontsCount(): number {
    return this.fonts.length;
  }

  getFillsCount(): number {
    return this.fills.length;
  }

  getBordersCount(): number {
    return this.borders.length;
  }

  getCellXfsCount(): number {
    return this.cellXfs.length;
  }
}
