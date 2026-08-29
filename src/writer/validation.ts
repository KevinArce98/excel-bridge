import { CellValidation, DataValidationOperator } from '../core/types';
import { dateToExcelSerial } from '../core/date-utils';

type RangeOperator = Exclude<DataValidationOperator, never>;

/**
 * Typed builders for cell data validations. Each returns a {@link CellValidation}
 * ready to drop into a sheet's `validations` array or `Workbook.addValidation`.
 */
export const dataValidation = {
  list(range: string, values: string[]): CellValidation {
    return { range, type: 'list', options: values.join(',') };
  },

  wholeNumber(
    range: string,
    operator: RangeOperator,
    value: number,
    value2?: number
  ): CellValidation {
    return numericRule('whole', range, operator, value, value2);
  },

  decimal(range: string, operator: RangeOperator, value: number, value2?: number): CellValidation {
    return numericRule('decimal', range, operator, value, value2);
  },

  textLength(
    range: string,
    operator: RangeOperator,
    value: number,
    value2?: number
  ): CellValidation {
    return numericRule('textLength', range, operator, value, value2);
  },

  dateBetween(range: string, start: Date, end: Date): CellValidation {
    return {
      range,
      type: 'date',
      operator: 'between',
      formula1: String(dateToExcelSerial(start)),
      formula2: String(dateToExcelSerial(end)),
      options: '',
    };
  },
};

const numericRule = (
  type: CellValidation['type'],
  range: string,
  operator: RangeOperator,
  value: number,
  value2?: number
): CellValidation => ({
  range,
  type,
  operator,
  formula1: String(value),
  ...(value2 !== undefined ? { formula2: String(value2) } : {}),
  options: '',
});
