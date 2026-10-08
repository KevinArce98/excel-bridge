import { invalidInput } from './errors';
import { BORDER_SIDES, BORDER_STYLES } from './borders';
import type { BorderLine, BorderSide, BorderSides } from './borders';
import { isDateFormatCode } from './date-utils';
import { CellStyle, ConditionalFormatStyle } from './types';

export function normalizeColor(color: string): string {
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(color);
  if (!match) {
    throw invalidInput(`Invalid colour "${color}": use #RGB, #RRGGBB or #AARRGGBB`);
  }
  let normalized = match[1];

  if (normalized.length === 3) {
    normalized = normalized
      .split('')
      .map(c => c + c)
      .join('');
  }

  if (normalized.length === 6) {
    normalized = 'FF' + normalized;
  }

  return normalized.toUpperCase();
}

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

type Table = Map<string, number>;

const FIRST_CUSTOM_NUM_FMT_ID = 164;
const BUILT_IN_DATE_NUM_FMT_ID = 14;

type ApplyFlags = Partial<
  Record<
    'applyFont' | 'applyFill' | 'applyBorder' | 'applyNumberFormat' | 'applyAlignment',
    unknown
  >
>;

const KEYED_STYLE_FIELDS: Record<keyof CellStyle, true> = {
  background: true,
  border: true,
  bold: true,
  italic: true,
  underline: true,
  color: true,
  fontSize: true,
  fontName: true,
  align: true,
  verticalAlign: true,
  wrapText: true,
  numberFormat: true,
};

const STYLE_FIELDS = Object.keys(KEYED_STYLE_FIELDS) as (keyof CellStyle)[];

const intern = (table: Table, xml: string): number => {
  let id = table.get(xml);
  if (id === undefined) {
    id = table.size;
    table.set(xml, id);
  }
  return id;
};

const joinEntries = (table: Table): string => Array.from(table.keys()).join('\n');

const escapeAttr = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const indentedLines = (...lines: unknown[]): string =>
  lines
    .filter(Boolean)
    .map(line => `\n      ${line}`)
    .join('');

const fontXml = ({ bold, italic, underline, color, fontSize, fontName }: CellStyle): string =>
  '    <font>' +
  indentedLines(
    bold && '<b/>',
    italic && '<i/>',
    underline && '<u/>',
    `<sz val="${fontSize || 11}"/>`,
    color && `<color rgb="${normalizeColor(color)}"/>`,
    `<name val="${escapeAttr(fontName || 'Calibri')}"/>`
  ) +
  '\n    </font>';

const fillXml = (background?: string): string =>
  `    <fill>\n      <patternFill patternType="${background ? 'solid' : 'none'}"${
    background
      ? `>\n        <fgColor rgb="${normalizeColor(background)}"/>\n      </patternFill>`
      : '/>'
  }\n    </fill>`;

const OPAQUE_BLACK = 'FF000000';

const borderSideXml = (side: string, spec?: BorderSide): string => {
  if (!spec) return `<${side}/>`;
  const { style, color }: Partial<BorderLine> = typeof spec === 'string' ? { style: spec } : spec;
  if (!style || !BORDER_STYLES.includes(style)) {
    throw invalidInput(
      `Invalid border style "${style ?? JSON.stringify(spec)}": use ${BORDER_STYLES.join(', ')}`
    );
  }
  return `<${side} style="${style}"><color rgb="${color ? normalizeColor(color) : OPAQUE_BLACK}"/></${side}>`;
};

const isPerSide = (border: BorderSide | BorderSides): border is BorderSides =>
  typeof border === 'object' && border.style === undefined && border.color === undefined;

const borderXml = (border: CellStyle['border'] = false): string => {
  const whole = border === true ? 'thin' : border;
  const sides: BorderSides = !whole
    ? {}
    : isPerSide(whole)
      ? whole
      : { left: whole, right: whole, top: whole, bottom: whole };
  return (
    '    <border>' +
    BORDER_SIDES.map(side => borderSideXml(side, sides[side])).join('') +
    '\n      <diagonal/>\n    </border>'
  );
};

const alignmentXml = ({ align, verticalAlign, wrapText }: CellStyle): string =>
  align || verticalAlign || wrapText
    ? `<alignment${align ? ` horizontal="${escapeAttr(align)}"` : ''}${
        verticalAlign
          ? ` vertical="${escapeAttr(verticalAlign === 'middle' ? 'center' : verticalAlign)}"`
          : ''
      }${wrapText ? ' wrapText="1"' : ''}/>`
    : '';

const dxfXml = ({ bold, italic, color, background }: ConditionalFormatStyle): string => {
  const font =
    bold || italic || color
      ? `<font>${bold ? '<b/>' : ''}${italic ? '<i/>' : ''}${color ? `<color rgb="${normalizeColor(color)}"/>` : ''}</font>`
      : '';
  const fill = background
    ? `<fill><patternFill><bgColor rgb="${normalizeColor(background)}"/></patternFill></fill>`
    : '';
  return `    <dxf>${font}${fill}</dxf>`;
};

const styleKey = (style: CellStyle): string =>
  JSON.stringify(STYLE_FIELDS.map(field => style[field]));

export class StyleManager {
  private fonts: Table = new Map();
  private fills: Table = new Map();
  private borders: Table = new Map();
  private numFmts: Table = new Map();
  private dxfs: Table = new Map();
  private cellXfs: Table = new Map();
  private styleIds: Map<string, number> = new Map();

  constructor() {
    intern(this.fonts, fontXml({}));
    intern(this.fills, fillXml());
    intern(this.fills, '    <fill>\n      <patternFill patternType="gray125"/>\n    </fill>');
    intern(this.borders, borderXml());
    this.internXf(0, 0, 0, 0);
  }

  getStyleId(style: CellStyle): number {
    const key = styleKey(style);
    let id = this.styleIds.get(key);
    if (id === undefined) {
      id = this.registerStyle(style);
      this.styleIds.set(key, id);
    }
    return id;
  }

  getDateStyleId(style?: CellStyle): number {
    const key = style ? `date${styleKey(style)}` : 'date';
    let id = this.styleIds.get(key);
    if (id === undefined) {
      id = this.registerStyle(style ?? {}, BUILT_IN_DATE_NUM_FMT_ID);
      this.styleIds.set(key, id);
    }
    return id;
  }

  private registerStyle(style: CellStyle, defaultNumFmtId = 0): number {
    const { background, bold, italic, underline, color, fontSize, fontName } = style;
    const numberFormat =
      defaultNumFmtId && style.numberFormat && !isDateFormatCode(style.numberFormat)
        ? undefined
        : style.numberFormat;
    const alignment = alignmentXml(style);
    const fontId = intern(this.fonts, fontXml(style));
    const fillId = intern(this.fills, fillXml(background));
    const borderId = intern(this.borders, borderXml(style.border));
    const numFmtId = numberFormat
      ? FIRST_CUSTOM_NUM_FMT_ID + intern(this.numFmts, numberFormat)
      : defaultNumFmtId;
    return this.internXf(
      fontId,
      fillId,
      borderId,
      numFmtId,
      {
        applyFont: bold || italic || underline || color || fontSize || fontName,
        applyFill: background,
        applyBorder: borderId,
        applyNumberFormat: numFmtId,
        applyAlignment: alignment,
      },
      alignment
    );
  }

  private internXf(
    fontId: number,
    fillId: number,
    borderId: number,
    numFmtId: number,
    applied: ApplyFlags = {},
    alignment = ''
  ): number {
    const flags = Object.entries(applied)
      .filter(([, enabled]) => enabled)
      .map(([flag]) => ` ${flag}="1"`)
      .join('');
    const open = `    <xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${flags}`;
    return intern(
      this.cellXfs,
      alignment ? `${open}>\n      ${alignment}\n    </xf>` : `${open}/>`
    );
  }

  generateFontsXml(): string {
    return joinEntries(this.fonts);
  }

  generateFillsXml(): string {
    return joinEntries(this.fills);
  }

  generateBordersXml(): string {
    return joinEntries(this.borders);
  }

  generateCellXfsXml(): string {
    return joinEntries(this.cellXfs);
  }

  generateNumFmtsXml(): string {
    return Array.from(
      this.numFmts.keys(),
      (code, index) =>
        `    <numFmt numFmtId="${FIRST_CUSTOM_NUM_FMT_ID + index}" formatCode="${escapeAttr(code)}"/>`
    ).join('\n');
  }

  getNumFmtsCount(): number {
    return this.numFmts.size;
  }

  getDxfId(style: ConditionalFormatStyle): number {
    return intern(this.dxfs, dxfXml(style));
  }

  generateDxfsXml(): string {
    return joinEntries(this.dxfs);
  }

  getDxfsCount(): number {
    return this.dxfs.size;
  }

  getFontsCount(): number {
    return this.fonts.size;
  }

  getFillsCount(): number {
    return this.fills.size;
  }

  getBordersCount(): number {
    return this.borders.size;
  }

  getCellXfsCount(): number {
    return this.cellXfs.size;
  }
}
