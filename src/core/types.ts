import type { EXCEL_ERRORS } from './cells';
import type { CellBorder } from './borders';

export type {
  CellBorder,
  BorderSide,
  BorderSides,
  BorderLine,
  BorderLines,
  BorderSideName,
  BorderStyleName,
  ParsedBorder,
} from './borders';

export type ExcelErrorValue = (typeof EXCEL_ERRORS)[number];

export interface ErrorCell {
  error: ExcelErrorValue;
  formula?: never;
  text?: never;
}

export interface TextCell {
  text: string;
  formula?: never;
  error?: never;
}

export type FormulaResult = string | number | boolean | Date | ErrorCell;

export interface FormulaCell {
  formula: string;
  result?: FormulaResult;
  text?: never;
  error?: never;
}

export type CellValue =
  string | number | boolean | Date | null | undefined | FormulaCell | TextCell | ErrorCell;

export type DataValidationType =
  'list' | 'whole' | 'decimal' | 'textLength' | 'date' | 'time' | 'custom';

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

export interface CellValidation {
  range: string;
  type?: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
}

export interface ConditionalFormatStyle {
  background?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

export type ConditionalFormatOperator =
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'between'
  | 'notBetween';

export interface CellValueConditionalFormat {
  type: 'cellValue';
  range: string;
  operator: ConditionalFormatOperator;
  value: number | string;
  value2?: number | string;
  style: ConditionalFormatStyle;
}

export interface ExpressionConditionalFormat {
  type: 'expression';
  range: string;
  formula: string;
  style: ConditionalFormatStyle;
}

export interface ColorScaleConditionalFormat {
  type: 'colorScale';
  range: string;
  colors: [string, string] | [string, string, string];
}

export type ConditionalFormat =
  CellValueConditionalFormat | ExpressionConditionalFormat | ColorScaleConditionalFormat;

export interface AutoFilter {
  range: string;
}

export interface ExternalHyperlink {
  range: string;
  url: string;
  location?: never;
  tooltip?: string;
  display?: string;
}

export interface InternalHyperlink {
  range: string;
  location: string;
  url?: never;
  tooltip?: string;
  display?: string;
}

export type Hyperlink = ExternalHyperlink | InternalHyperlink;

export interface SheetLayout {
  freezePane?: { row?: number; col?: number };
  columnWidths?: number[];
  rowHeights?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
}

export interface CellStyle<Border extends CellBorder = CellBorder> {
  background?: string;
  border?: Border;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSize?: number;
  fontName?: string;
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
}
