/** A value that can be written into a worksheet cell. */
export type CellValue = string | number | boolean | Date | null | undefined;

/** A list data-validation applied to a cell range. */
export interface CellValidation {
  range: string;
  options: string;
}

/** Visual and number-format options for a single cell. */
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
  /** Custom Excel number-format code, e.g. "0.00" or "#,##0". */
  numberFormat?: string;
}
