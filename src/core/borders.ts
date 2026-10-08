export const BORDER_STYLES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
] as const;

export const BORDER_SIDES = ['left', 'right', 'top', 'bottom'] as const;

export type BorderStyleName = (typeof BORDER_STYLES)[number];

export type BorderSideName = (typeof BORDER_SIDES)[number];

export interface BorderLine {
  style: BorderStyleName;
  color?: string;
}

export type BorderSide = BorderStyleName | BorderLine;

export type BorderLines = { [side in BorderSideName]?: BorderLine };

export type BorderSides = { [side in BorderSideName]?: BorderSide } & {
  style?: never;
  color?: never;
};

export type ParsedBorder = true | BorderLines;

export type CellBorder = boolean | BorderSide | BorderSides;
