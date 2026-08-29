export type CellValue = string | number | boolean | Date | null | undefined;

export type DataValidationType = 'list' | 'whole' | 'decimal' | 'textLength' | 'date';

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
  /** Legacy list shorthand: comma-separated allowed values (used when `type` is 'list' or unset). */
  options: string;
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

export interface CellStyle {
  background?: string;
  border?: boolean;
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
