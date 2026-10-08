import { EXCEL_LIMITS } from './date-utils';
import { MAX_ROW_HEIGHT, isRowHeight } from './row-height';
import { rowIndexes } from './rows';
import type { SheetLayout } from './types';

export interface PreparedLayout {
  rows: number[];
  hiddenColumns: Set<number>;
  rowAttributes: (rowIndex: number) => string;
}

const checkIndex = (kind: string, index: number, limit: number): void => {
  if (!Number.isInteger(index) || index < 0 || index >= limit) {
    throw new Error(`${kind} index ${index} must be a whole number from 0 to ${limit - 1}`);
  }
};

export const prepareLayout = ({
  rowHeights = {},
  hiddenRows = [],
  hiddenColumns = [],
}: SheetLayout): PreparedLayout => {
  const heightRows = rowIndexes(rowHeights);
  const hiddenRowSet = new Set(hiddenRows);
  const hiddenColumnSet = new Set(hiddenColumns);

  heightRows.forEach(row => {
    checkIndex('Row', row, EXCEL_LIMITS.MAX_ROWS);
    if (!isRowHeight(rowHeights[row])) {
      throw new Error(
        `Row ${row + 1} height ${rowHeights[row]} must be above 0 and at most ${MAX_ROW_HEIGHT} points`
      );
    }
  });
  hiddenRowSet.forEach(row => checkIndex('Row', row, EXCEL_LIMITS.MAX_ROWS));
  hiddenColumnSet.forEach(column => checkIndex('Column', column, EXCEL_LIMITS.MAX_COLS));

  return {
    rows: [...new Set([...heightRows, ...hiddenRowSet])].sort((a, b) => a - b),
    hiddenColumns: hiddenColumnSet,
    rowAttributes: row =>
      (rowHeights[row] ? ` ht="${rowHeights[row]}" customHeight="1"` : '') +
      (hiddenRowSet.has(row) ? ' hidden="1"' : ''),
  };
};
