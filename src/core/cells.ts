import { CellValue, ExcelErrorValue, FormulaResult } from './types';

export const EXCEL_ERRORS = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
] as const;

export const isExcelErrorValue = (value: unknown): value is ExcelErrorValue =>
  EXCEL_ERRORS.includes(value as ExcelErrorValue);

export interface CellParts {
  formula?: string;
  value: FormulaResult | null | undefined;
}

export const splitCell = (cell: CellValue): CellParts => {
  if (typeof cell === 'object' && cell !== null && !(cell instanceof Date)) {
    if (typeof cell.formula === 'string') {
      return { formula: cell.formula.replace(/^=/, ''), value: cell.result };
    }
    if (typeof cell.text === 'string') return { value: cell.text };
  }
  return { value: cell as FormulaResult | null | undefined };
};
